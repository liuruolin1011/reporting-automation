import os
from data_processing import process_data
from flask import Flask, request, render_template, send_file, abort

dirpath_dst = r'C:\Users\RUOLINLIU\Desktop\Heatmap Automation Flask'

app = Flask(__name__)

# Limit the access to specific IPs
allowed_ips = ['22.232.100.153',  # Ruolin
               '22.232.100.33',   # Brad
               '21.232.104.75']   # Apoorv

@app.before_request
def limit_access():
    if request.remote_addr not in allowed_ips:
        abort(403)

@app.route('/', methods=['GET', 'POST'])
def date_input():
    available_years = ['2021', '2022', '2023', '2024']
    available_months = {
        '01': 'January', '02': 'February', '03': 'March', '04': 'April',
        '05': 'May', '06': 'June', '07': 'July', '08': 'August',
        '09': 'September', '10': 'October', '11': 'November', '12': 'December'
    }

    if request.method == 'POST':
        selected_year = request.form.get('year')
        selected_month = request.form.get('month')
        if not selected_year or not selected_month:
            return "Please select both a year and month", 400  # Better error handling
        return f"Selected Year: {selected_year}, Selected Month: {available_months.get(selected_month, 'Invalid month')}"

    return render_template('date_input.html', available_years=available_years, available_months=available_months)


@app.route('/completed', methods=['POST'])
def index():
    previous_month = request.form.get('previous_month')
    previous_year = request.form.get('previous_year')
    current_month = request.form.get('current_month')
    current_year = request.form.get('current_year')

    # Add error handling for missing data
    if not previous_month or not previous_year or not current_month or not current_year:
        return "All date fields are required", 400

    input_folder_path = f"data_input/{previous_month} {previous_year}"
    output_folder_path = f"data_output/{current_month} {current_year}"

    try:
        filename, removal_list, current_restricted_cust, current_active_cust, additional_cust, negative_news, total_offboard_cust = process_data(
            previous_month=previous_month, previous_year=previous_year,
            current_month=current_month, current_year=current_year,
            input_folder_path=input_folder_path, output_folder_path=output_folder_path
        )
    except Exception as e:
        return f"Error during processing: {str(e)}", 500  # Handle processing errors gracefully

    parameters = {
        f'{current_month} Active Customers:': len(current_active_cust.index),
        f'{current_month} Restricted Customers:': len(current_restricted_cust.index),
        f'{current_month} Offboard Customers:': len(total_offboard_cust.index),
        f'Customers added to {current_month} Active list:': len(additional_cust.index),
        f'Customers removed from {previous_month} Active list:': len(removal_list.index),
        f'{current_month} Negative News:': len(negative_news)
    }

    return render_template('result.html', ext_filename=filename, parameters=parameters)


@app.route('/download/<filename>')
def download_excel(filename):
    filepath = os.path.join(dirpath_dst, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    else:
        return f"File '{filename}' not found.", 404


if __name__ == '__main__':
    app.run(debug=True, host='22.232.100.153', port=5014)  # Ensure this IP is accessible to the network
