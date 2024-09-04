from flask import Flask, render_template, request, send_file, send_from_directory
import pandas as pd
import io

app = Flask(__name__, static_folder='static', template_folder='templates')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Retrieve the form data
        form_data = request.form.to_dict()

        # Generate a new Excel file from the form data
        excel_file = generate_excel(form_data)
        
        # Send the Excel file as a response
        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Data_Curah_Hujan_Harian.xlsx'
        )

    return render_template('index.html')

def generate_excel(data):
    output = io.BytesIO()

    # Create a new Excel file with XlsxWriter
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Data Curah Hujan Harian')

        # Define some formats
        title_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 16
        })
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#D9EAD3',
            'border': 1
        })
        normal_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        yellow_fill = workbook.add_format({
            'fg_color': '#FFFF00',
            'border': 1
        })
        green_fill = workbook.add_format({
            'fg_color': '#00FF00',
            'border': 1
        })

        # Set column widths
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:M', 10)

        # Write the title
        worksheet.merge_range('A1:M1', 'DATA CURAH HUJAN HARIAN', title_format)
        worksheet.merge_range('A2:M2', f'Tahun : {data.get("tahun", "[Tahun]")}', normal_format)

        # Write station information
        worksheet.merge_range('A4:B4', 'Nama Stasiun', header_format)
        worksheet.merge_range('C4:D4', data.get('nama_stasiun', '[Nama Stasiun]'), normal_format)

        worksheet.merge_range('A5:B5', 'Kode Stasiun', header_format)
        worksheet.merge_range('C5:D5', data.get('kode_stasiun', '[Kode Stasiun]'), normal_format)

        worksheet.merge_range('F5:G5', 'Wilayah Sungai', header_format)
        worksheet.merge_range('H5:I5', data.get('wilayah_sungai', '[Wilayah Sungai]'), normal_format)

        worksheet.merge_range('F6:G6', 'Kelurahan/Desa', header_format)
        worksheet.merge_range('H6:I6', data.get('kelurahan', '[Kelurahan/Desa]'), normal_format)

        worksheet.merge_range('A6:B6', 'Lintang Selatan', header_format)
        worksheet.merge_range('C6:D6', data.get('longitude', '[Longitude]'), normal_format)

        worksheet.merge_range('F7:G7', 'Kecamatan', header_format)
        worksheet.merge_range('H7:I7', data.get('kecamatan', '[Kecamatan]'), normal_format)

        worksheet.merge_range('F8:G8', 'Kabupaten', header_format)
        worksheet.merge_range('H8:I8', data.get('kabupaten', '[Kabupaten]'), normal_format)

        worksheet.merge_range('A7:B7', 'Bujur Timur', header_format)
        worksheet.merge_range('C7:D7', data.get('latitude', '[Latitude]'), normal_format)

        worksheet.merge_range('A8:B8', 'Elevasi', header_format)
        worksheet.merge_range('C8:D8', data.get('elevation', '[Elevation]'), normal_format)

        worksheet.merge_range('J5:K5', 'Kode Database', header_format)
        worksheet.merge_range('L5:M5', data.get('kode_database', '[Kode Database]'), normal_format)

        worksheet.merge_range('J6:K6', 'Tahun Pendirian', header_format)
        worksheet.merge_range('L6:M6', data.get('tahun_pendirian', '[Tahun Pendirian]'), normal_format)

        worksheet.merge_range('J7:K7', 'Tipe Alat', header_format)
        worksheet.merge_range('L7:M7', data.get('tipe_alat', '[Tipe Alat]'), normal_format)

        worksheet.merge_range('J8:K8', 'Pengelola', header_format)
        worksheet.merge_range('L8:M8', data.get('pengelola', '[Pengelola]'), normal_format)

        # Write rainfall data table headers
        worksheet.merge_range('A10:A11', 'Tanggal', header_format)
        worksheet.merge_range('B10:M10', 'Bulan', header_format)
        for month in range(1, 13):
            worksheet.write(10, month, f'{month}', header_format)

        # Write daily rainfall data
        for day in range(1, 32):
            worksheet.write(day + 10, 0, day, normal_format)
            for month in range(1, 13):
                worksheet.write(day + 10, month, data.get(f'day{day}_month{month}', 0), normal_format)

        # Write totals
        totals = ['Total', 'Periode1', 'Periode2', 'Periode3', 'Maksimum', 'DataHujan']
        for i, total in enumerate(totals):
            worksheet.write(42 + i, 0, total, header_format)
            for month in range(1, 13):
                worksheet.write(42 + i, month, data.get(f'{total.lower()}_month{month}', 0), normal_format)
                # worksheet.write(42 + i, 13, data.get(f'datahujan_month{{ month }}', 0), normal_format)
        # Write analysis table headers
        analysis_start = 49
        analysis_headers = ['No', 'Bulan', 'Curah Hujan', 'Sk*', '[Sk*]', 'Dy^2', 'Dy', 'Sk**', '[Sk**]']
        for i, header in enumerate(analysis_headers):
            worksheet.write(analysis_start, i, header, header_format)

        # Write analysis data
        months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']
        for i, month in enumerate(months):
            worksheet.write(analysis_start + i + 1, 0, i + 1, normal_format)
            worksheet.write(analysis_start + i + 1, 1, month, normal_format)
            worksheet.write(analysis_start + i + 1, 2, data.get(f'curah_hujan_{i}', 0), normal_format)
            worksheet.write(analysis_start + i + 1, 3, data.get(f'sk_{i}', 0), normal_format)
            worksheet.write(analysis_start + i + 1, 4, data.get(f'sk_brackets_{i}', 0), normal_format)
            worksheet.write(analysis_start + i + 1, 5, data.get(f'dy2_{i}', 0), normal_format)
            worksheet.write(analysis_start + i + 1, 6, data.get(f'dy_{i}', 0), normal_format)
            worksheet.write(analysis_start + i + 1, 7, data.get(f'sk_star_{i}', 0), normal_format)
            worksheet.write(analysis_start + i + 1, 8, data.get(f'sk_star_brackets_{i}', 0), normal_format)

        # Write final analysis summary
        worksheet.write(f'B{analysis_start + 14}', 'Rerata', header_format)
        worksheet.write(f'C{analysis_start + 14}', data.get('rerata_curah_hujan', 0), normal_format)
        
        worksheet.write(f'E{analysis_start + 14}', data.get('rerata_sk_brackets', 0), normal_format)

        worksheet.write(f'B{analysis_start + 15}', 'Jumlah', header_format)
        worksheet.write(f'C{analysis_start + 15}', data.get('jumlah_curah_hujan', 0), normal_format)
        
        worksheet.write(f'F{analysis_start + 15}', data.get('jumlah_dy2', 0), normal_format)
        
        worksheet.write(f'B{analysis_start + 16}', 'Maks', header_format)
        worksheet.write(f'C{analysis_start + 16}', data.get('maks_curah_hujan', 0), normal_format)
        
        worksheet.write(f'H{analysis_start + 16}', data.get('maks_sk', 0), normal_format)
        worksheet.write(f'I{analysis_start + 16}', data.get('maks_sk_brackets', 0), normal_format)
        
        worksheet.write(f'B{analysis_start + 17}', 'Min', header_format)
        worksheet.write(f'C{analysis_start + 17}', data.get('min_curah_hujan', 0), normal_format)
        
        worksheet.write(f'H{analysis_start + 17}', data.get('min_sk', 0), normal_format)
        worksheet.write(f'I{analysis_start + 17}', data.get('min_sk_brackets', 0), normal_format)

        worksheet.write(f'B{analysis_start + 19}', 'Hasil analisis :', normal_format)

        worksheet.write(f'B{analysis_start + 20}', 'n', header_format)
        worksheet.write(f'C{analysis_start + 20}', data.get('n_value', 12), normal_format)

        worksheet.write(f'B{analysis_start + 21}', 'Sk**mak', header_format)
        worksheet.write(f'C{analysis_start + 21}', data.get('sk_mak', 0), normal_format)

        worksheet.write(f'B{analysis_start + 22}', 'Sk**min', header_format)
        worksheet.write(f'C{analysis_start + 22}', data.get('sk_min', 0), normal_format)

        worksheet.write(f'B{analysis_start + 23}', 'Q = Sk**mak', header_format)
        worksheet.write(f'C{analysis_start + 23}', data.get('sk_mak', 0), normal_format)

        worksheet.write(f'B{analysis_start + 24}', 'R = Sk**mak - Sk**min', header_format)
        worksheet.write(f'C{analysis_start + 24}', data.get('r_sk_diff', 0), normal_format)

        worksheet.write(f'B{analysis_start + 25}', 'Q/n^0.5', header_format)
        worksheet.write(f'C{analysis_start + 25}', data.get('q_over_n', 0), normal_format)
        worksheet.write(f'D{analysis_start + 25}', '< dengan probabilitas 95% dari tabel', header_format)
        worksheet.write(f'E{analysis_start + 25}', data.get('q_value', 0), normal_format)
        if (data.get('q_over_n') < data.get('q_value')):
            worksheet.write(f'F{analysis_start + 25}', data.get('q_over_n_status_text', 'OK!'), green_fill)
        else:
            worksheet.write(f'F{analysis_start + 25}', data.get('q_over_n_status_text', 'NOT OK!'), yellow_fill)

        # worksheet.write(f'F{analysis_start + 25}', data.get('q_over_n_status_text', '-'), normal_format)

        worksheet.write(f'B{analysis_start + 26}', 'R/n^0.5', header_format)
        worksheet.write(f'C{analysis_start + 26}', data.get('r_over_n', 0), normal_format)
        worksheet.write(f'D{analysis_start + 26}', '< dengan probabilitas 95% dari tabel', header_format)
        worksheet.write(f'E{analysis_start + 26}', data.get('r_value', 0), normal_format)
        if (data.get('r_over_n') < data.get('r_value')):
            worksheet.write(f'F{analysis_start + 26}', data.get('r_over_n_status_text', 'OK!'), green_fill)
        else:
            worksheet.write(f'F{analysis_start + 26}', data.get('r_over_n_status_text', 'NOT OK!'), yellow_fill)

        # worksheet.write(f'F{analysis_start + 26}', data.get('r_over_n_status_text', '-'), normal_format)

        # Write the final analysis table for Q/n^0.5 and R/n^0.5 values
        final_table_start = analysis_start + 28
        worksheet.merge_range(f'B{final_table_start}:D{final_table_start}', 'Nilai Q/n^0.5 dan R/n^0.5', normal_format)

        final_table_headers = ['n', 'Q/n^0.5', 'Q/n^0.5', 'Q/n^0.5', 'R/n^0.5', 'R/n^0.5', 'R/n^0.5']
        for i, header in enumerate(final_table_headers):
            worksheet.write(final_table_start, i + 1, header, header_format)

        # Sample values for Q/n^0.5 and R/n^0.5
        final_table_data = [
            (" ", "90%", "95%", "99%", "90%", "95%", "99%"),
            (10, 1.05, 1.14, 1.29, 1.21, 1.28, 1.38),
            (20, 1.10, 1.22, 1.42, 1.34, 1.43, 1.60),
            (30, 1.12, 1.24, 1.48, 1.40, 1.50, 1.70),
            (40, 1.14, 1.27, 1.52, 1.44, 1.55, 1.78),
            (100, 1.17, 1.29, 1.63, 1.62, 1.75, 2.00)
        ]

        for i, row in enumerate(final_table_data):
            for j, value in enumerate(row):
                worksheet.write(final_table_start + 1 + i, j + 1, value, normal_format)

        # Final source citation
        worksheet.merge_range(f'C{final_table_start + 9}:H{final_table_start + 9}', 'Sumber: Sri Harto, 1993: 168', normal_format)

                # Write "Uji Abnormalitas Data" Table
        abnormal_data_start = final_table_start + 11  # Start row for the new table
        
        # Write headers for the "Uji Abnormalitas Data" table
        worksheet.write(abnormal_data_start, 1, 'No', header_format)
        worksheet.write(abnormal_data_start, 2, 'Bulan', header_format)
        worksheet.write(abnormal_data_start, 3, 'Curah Hujan (mm)', header_format)
        worksheet.write(abnormal_data_start, 4, 'Log X', header_format)

        # List of months
        months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']

        # Write data rows for the "Uji Abnormalitas Data" table
        for i, month in enumerate(months):
            worksheet.write(abnormal_data_start + i + 1, 1, i + 1, normal_format)  # Write the row number
            worksheet.write(abnormal_data_start + i + 1, 2, month, normal_format)  # Write the month name
            worksheet.write(abnormal_data_start + i + 1, 3, data.get(f'curah_hujan_x_{i}', 0), normal_format)  # Write Curah Hujan (mm)
            worksheet.write(abnormal_data_start + i + 1, 4, data.get(f'logx_{i}', 0), normal_format)  # Write Log X

        # Write the additional rows for Stdev, Mean, Kn, Xh, and Xi
        worksheet.write(abnormal_data_start + 13, 2, 'Stdev', header_format)
        worksheet.write(abnormal_data_start + 13, 3, data.get('stdev', 0), normal_format)

        worksheet.write(abnormal_data_start + 14, 2, 'Mean', header_format)
        worksheet.write(abnormal_data_start + 14, 3, data.get('xmean', 0), normal_format)

        worksheet.write(abnormal_data_start + 15, 2, 'Kn', header_format)
        worksheet.write(abnormal_data_start + 15, 3, data.get('kn', 2.13), normal_format)

        worksheet.write(abnormal_data_start + 16, 2, 'Nilai Ambang Atas', header_format)

        worksheet.write(abnormal_data_start + 17, 2, 'Xh=', header_format)
        worksheet.write(abnormal_data_start + 17, 3, data.get('Xh', 0), normal_format)

        worksheet.write(abnormal_data_start + 18, 2, 'Nilai Ambang Bawah', header_format)

        worksheet.write(abnormal_data_start + 19, 2, 'Xi=', header_format)
        worksheet.write(abnormal_data_start + 19, 3, data.get('Xi', 0), normal_format)

        # Write the status of the test
        worksheet.write(abnormal_data_start + 20, 2, data.get('status_uji', '-'), normal_format)

    output.seek(0)  # Rewind the buffer
    return output

if __name__ == '__main__':
    # app.run(port=5001, debug=True)
    app.run(host='103.183.92.89', port=5001, debug=False)
