from flask import Flask, request, render_template, redirect, url_for, send_from_directory, flash
import os
import pandas as pd
from werkzeug.utils import secure_filename
import json
import requests
from datetime import datetime
import logging

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Required for flash messages
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['RESULTS_FOLDER'] = 'results'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

# Ensure the upload and results folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULTS_FOLDER'], exist_ok=True)

# Logging configuration
logging.basicConfig(filename='error_log.txt', level=logging.ERROR, format='%(asctime)s %(message)s')

# Helper functions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def read_family_mapping(json_file):
    with open(json_file, 'r', encoding='utf-8') as file:
        family_mapping = json.load(file)
    return family_mapping

def read_addresses_and_family_types_from_directory(directory_path, family_mapping):
    all_family_types = []
    unique_emails = set()

    for filename in os.listdir(directory_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(directory_path, filename)
            excel_file = pd.ExcelFile(file_path)
            for sheet_name in excel_file.sheet_names:
                sheet_df = excel_file.parse(sheet_name)
                if '黑客币址' in sheet_df.columns and '币类型（单位）' in sheet_df.columns and '后缀名称' in sheet_df.columns and '黑客联系邮箱' in sheet_df.columns:
                    valid_rows = sheet_df[['黑客币址', '币类型（单位）', '后缀名称', '黑客联系邮箱']].dropna(subset=['黑客币址', '币类型（单位）'])
                    for index, row in valid_rows.iterrows():
                        raw_family_type = row.get('后缀名称', '未知')
                        raw_family_type = str(raw_family_type)
                        family_type = family_mapping.get(raw_family_type, '未知')
                        if family_type == '未知':
                            print(f"Unmapped type: {raw_family_type}")
                        all_family_types.append(family_type)

                        hacker_email = row.get('黑客联系邮箱')
                        if pd.notna(hacker_email):
                            unique_emails.add(hacker_email)

    return all_family_types, unique_emails

def read_addresses_and_chains_from_directory(directory_path):
    all_addresses_and_chains = set()
    for filename in os.listdir(directory_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(directory_path, filename)
            excel_file = pd.ExcelFile(file_path)
            for sheet_name in excel_file.sheet_names:
                sheet_df = excel_file.parse(sheet_name)
                if '黑客币址' in sheet_df.columns and '币类型（单位）' in sheet_df.columns:
                    valid_rows = sheet_df[['黑客币址', '币类型（单位）']].dropna()
                    for index, row in valid_rows.iterrows():
                        address = row['黑客币址']
                        chain_short_name = row['币类型（单位）']
                        all_addresses_and_chains.add((address, chain_short_name))
    return all_addresses_and_chains

def convert_to_usd(amount, currency, exchange_rates):
    rate = exchange_rates.get(currency.upper(), 0)
    return amount * rate

def get_address_summary(address, chain_short_name, API_KEY):
    url = 'https://www.oklink.com/api/v5/explorer/address/address-summary'
    headers = {'Ok-Access-Key': API_KEY}
    params = {'chainShortName': chain_short_name, 'address': address}
    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    return response.json()

def parse_transactions(data, chain_short_name, exchange_rates):
    transaction_summary = {}
    for item in data:
        first_tx_time = int(item['firstTransactionTime']) / 1000
        last_tx_time = int(item['lastTransactionTime']) / 1000

        first_tx_date = datetime.utcfromtimestamp(first_tx_time).strftime('%Y-%m')
        last_tx_date = datetime.utcfromtimestamp(last_tx_time).strftime('%Y-%m')

        if last_tx_date not in transaction_summary:
            transaction_summary[last_tx_date] = {'receive': 0}
        transaction_summary[last_tx_date]['receive'] += convert_to_usd(float(item['receiveAmount']), chain_short_name, exchange_rates)
    return transaction_summary

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        files = request.files.getlist('files[]')
        if not files or any(f.filename == '' for f in files):
            flash('没有选择文件。')
            return redirect(request.url)

        saved_files = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                saved_files.append(file_path)
            else:
                flash('无效的文件类型。')
                return redirect(request.url)

        # Load family mapping from a fixed file
        json_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'family.json')
        family_mapping = read_family_mapping(json_file_path)

        # Process family data
        family_types, unique_emails = read_addresses_and_family_types_from_directory(app.config['UPLOAD_FOLDER'], family_mapping)

        # Calculate statistics
        family_type_counts = pd.Series(family_types).value_counts()
        family_type_counts_df = family_type_counts.reset_index()
        family_type_counts_df.columns = ['家族名称', '数量']
        family_type_counts_df['家族名称'] = family_type_counts_df.apply(lambda x: 'others' if x['数量'] < 6 else x['家族名称'], axis=1)
        family_type_counts_df = family_type_counts_df.groupby('家族名称').sum().reset_index()

        # Save results
        output_file = os.path.join(app.config['RESULTS_FOLDER'], 'ransomware_family_statistics.xlsx')
        with pd.ExcelWriter(output_file) as writer:
            family_type_counts_df.to_excel(writer, sheet_name='Ransomware Family Statistics', index=False)

        email_counts_df = pd.DataFrame({'黑客联系邮箱数量': [len(unique_emails)]})
        email_output_file = os.path.join(app.config['RESULTS_FOLDER'], 'hacker_email_statistics.xlsx')
        with pd.ExcelWriter(email_output_file) as writer:
            email_counts_df.to_excel(writer, sheet_name='Hacker Email Statistics', index=False)

        # Process address data
        all_addresses_and_chains = read_addresses_and_chains_from_directory(app.config['UPLOAD_FOLDER'])
        API_KEY = 'd7207c27-1a1d-4cb3-b672-2364d0b8748d'
        exchange_rates = {
            'BTC': 68000,
            'USDT': 1,
            # 可以根据需要添加其他币种的汇率
        }

        address_counts_by_currency = {}
        address_incoming = {}
        monthly_totals = {}

        for address, chain_short_name in all_addresses_and_chains:
            if chain_short_name not in address_counts_by_currency:
                address_counts_by_currency[chain_short_name] = 0
            address_counts_by_currency[chain_short_name] += 1

            summary = get_address_summary(address, chain_short_name, API_KEY)
            if not summary or 'data' not in summary:
                logging.error(f"Invalid summary data for address {address} on {chain_short_name}")
                continue

            transactions = summary['data']
            monthly_summary = parse_transactions(transactions, chain_short_name, exchange_rates)

            if address not in address_incoming:
                address_incoming[address] = (sum(data['receive'] for data in monthly_summary.values()), chain_short_name)

            for month, data in monthly_summary.items():
                if month not in monthly_totals:
                    monthly_totals[month] = {}
                if chain_short_name not in monthly_totals[month]:
                    monthly_totals[month][chain_short_name] = {'receive': 0}
                monthly_totals[month][chain_short_name]['receive'] += data['receive']

        # Save results
        address_counts_df = pd.DataFrame(list(address_counts_by_currency.items()), columns=['币种', '地址数量'])
        address_counts_df.to_csv(os.path.join(app.config['RESULTS_FOLDER'], 'address_counts_by_currency.csv'), index=False)

        top_20_addresses = sorted(address_incoming.items(), key=lambda x: x[1][0], reverse=True)[:20]
        top_20_addresses_usd = [(addr, incoming[0]) for addr, incoming in top_20_addresses]
        top_20_df = pd.DataFrame(top_20_addresses_usd, columns=['地址', '入金金额（美元）'])
        top_20_df.to_csv(os.path.join(app.config['RESULTS_FOLDER'], 'top_20_addresses.csv'), index=False)

        monthly_totals_usd = {}
        for month, chains in monthly_totals.items():
            if month not in monthly_totals_usd:
                monthly_totals_usd[month] = {}
            for chain, data in chains.items():
                if 'receive' not in monthly_totals_usd[month]:
                    monthly_totals_usd[month]['receive'] = 0
                monthly_totals_usd[month]['receive'] += data['receive']

        monthly_totals_df = pd.DataFrame(monthly_totals_usd).transpose().sort_index()
        monthly_totals_df.to_csv(os.path.join(app.config['RESULTS_FOLDER'], 'monthly_totals.csv'))

        flash('文件处理完成，结果已保存。')
        return redirect(url_for('upload_file'))

    return render_template('upload.html')

if __name__ == "__main__":
    app.run(debug=True)
