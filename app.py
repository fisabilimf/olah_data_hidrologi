from flask import Flask, render_template, request, send_file
import pandas as pd
import io

app = Flask(__name__)

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

        # Set column widths
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:N', 10)

        # Write the title
        worksheet.merge_range('A1:N1', 'DATA CURAH HUJAN HARIAN', title_format)
        worksheet.merge_range('A2:N2', f'Tahun : {data.get("tahun", "[Tahun]")}', normal_format)

        # Write station information
        worksheet.merge_range('A4:B4', 'NAMA STASIUN', header_format)
        worksheet.merge_range('C4:E4', data.get('nama_stasiun', '[Nama Stasiun]'), normal_format)

        worksheet.merge_range('A5:B5', 'Kode Stasiun', header_format)
        worksheet.merge_range('C5:E5', data.get('kode_stasiun', '[Kode Stasiun]'), normal_format)

        worksheet.merge_range('F4:G4', 'Wilayah Sungai', header_format)
        worksheet.merge_range('H4:J4', data.get('wilayah_sungai', '[Wilayah Sungai]'), normal_format)

        worksheet.merge_range('F5:G5', 'Kelurahan/Desa', header_format)
        worksheet.merge_range('H5:J5', data.get('kelurahan', '[Kelurahan/Desa]'), normal_format)

        worksheet.merge_range('A6:B6', 'Lintang Selatan', header_format)
        worksheet.merge_range('C6:E6', data.get('longitude', '[Longitude]'), normal_format)

        worksheet.merge_range('F6:G6', 'Kecamatan', header_format)
        worksheet.merge_range('H6:J6', data.get('kecamatan', '[Kecamatan]'), normal_format)

        worksheet.merge_range('A7:B7', 'Bujur Timur', header_format)
        worksheet.merge_range('C7:E7', data.get('latitude', '[Latitude]'), normal_format)

        worksheet.merge_range('A8:B8', 'Elevasi', header_format)
        worksheet.merge_range('C8:E8', data.get('elevation', '[Elevation]'), normal_format)

        # Write rainfall data table headers
        worksheet.write('A10', 'Tanggal', header_format)
        for month in range(1, 13):
            worksheet.write(9, month, f'Bulan {month}', header_format)

        # Write daily rainfall data
        for day in range(1, 32):
            worksheet.write(day + 9, 0, day, normal_format)
            for month in range(1, 13):
                worksheet.write(day + 9, month, data.get(f'day{day}_month{month}', 0), normal_format)

        # Write totals
        totals = ['Total', 'Periode1', 'Periode2', 'Periode3', 'Maksimum', 'Data Hujan']
        for i, total in enumerate(totals):
            worksheet.write(41 + i, 0, total, header_format)
            for month in range(1, 13):
                worksheet.write(41 + i, month, data.get(f'{total.lower()}_month{month}', 0), normal_format)

        # Write analysis table headers
        analysis_start = 48
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
        worksheet.write(f'A{analysis_start + 13}', 'Rerata', header_format)
        worksheet.write(f'C{analysis_start + 13}', data.get('rerata_curah_hujan', 0), normal_format)

        worksheet.write(f'A{analysis_start + 14}', 'Jumlah', header_format)
        worksheet.write(f'C{analysis_start + 14}', data.get('jumlah_curah_hujan', 0), normal_format)

        worksheet.write(f'A{analysis_start + 16}', 'Hasil analisis :', header_format)

        worksheet.write(f'A{analysis_start + 17}', 'n', header_format)
        worksheet.write(f'C{analysis_start + 17}', data.get('n', 12), yellow_fill)

        worksheet.write(f'A{analysis_start + 18}', 'Sk**mak', header_format)
        worksheet.write(f'C{analysis_start + 18}', data.get('sk_mak', 0), normal_format)

        worksheet.write(f'A{analysis_start + 19}', 'Sk**min', header_format)
        worksheet.write(f'C{analysis_start + 19}', data.get('sk_min', 0), normal_format)

        worksheet.write(f'A{analysis_start + 20}', 'Q = Sk**mak', header_format)
        worksheet.write(f'C{analysis_start + 20}', data.get('sk_mak', 0), normal_format)

        worksheet.write(f'A{analysis_start + 21}', 'R = Sk**mak - Sk**min', header_format)
        worksheet.write(f'C{analysis_start + 21}', data.get('r_sk_diff', 0), normal_format)

        worksheet.write(f'A{analysis_start + 22}', 'Q/n^0.5', header_format)
        worksheet.write(f'C{analysis_start + 22}', data.get('q_over_n', 0), normal_format)
        worksheet.write(f'D{analysis_start + 22}', '< dengan probabilitas 95% dari tabel', header_format)
        worksheet.write(f'E{analysis_start + 22}', data.get('q_table_value', 0), yellow_fill)
        worksheet.write(f'F{analysis_start + 22}', 'OK!!!', header_format)

        worksheet.write(f'A{analysis_start + 23}', 'R/n^0.5', header_format)
        worksheet.write(f'C{analysis_start + 23}', data.get('r_over_n', 0), normal_format)
        worksheet.write(f'D{analysis_start + 23}', '< dengan probabilitas 95% dari tabel', header_format)
        worksheet.write(f'E{analysis_start + 23}', data.get('r_table_value', 0), yellow_fill)
        worksheet.write(f'F{analysis_start + 23}', 'OK!!!', header_format)

        # Write the final analysis table for Q/n^0.5 and R/n^0.5 values
        final_table_start = analysis_start + 25
        worksheet.merge_range(f'C{final_table_start}:E{final_table_start}', 'Nilai Q/n^0.5 dan R/n^0.5', title_format)

        final_table_headers = ['n', 'Q/n^0.5 (90%)', 'Q/n^0.5 (95%)', 'Q/n^0.5 (99%)', 'R/n^0.5 (90%)', 'R/n^0.5 (95%)', 'R/n^0.5 (99%)']
        for i, header in enumerate(final_table_headers):
            worksheet.write(final_table_start + 1, i + 1, header, header_format)

        # Sample values for Q/n^0.5 and R/n^0.5
        final_table_data = [
            (10, 1.05, 1.14, 1.29, 1.21, 1.28, 1.38),
            (20, 1.10, 1.22, 1.42, 1.34, 1.43, 1.60),
            (30, 1.12, 1.24, 1.48, 1.40, 1.50, 1.70),
            (40, 1.14, 1.27, 1.52, 1.44, 1.55, 1.78),
            (100, 1.17, 1.29, 1.63, 1.62, 1.75, 2.00),
            (12, data.get('q_90_12', 1.05), data.get('q_95_12', 1.14), data.get('q_99_12', 1.29), 
                data.get('r_90_12', 1.21), data.get('r_95_12', 1.28), data.get('r_99_12', 1.38))
        ]

        for i, row in enumerate(final_table_data):
            for j, value in enumerate(row):
                worksheet.write(final_table_start + 2 + i, j + 1, value, normal_format)

        # Final source citation
        worksheet.merge_range(f'C{final_table_start + 9}:H{final_table_start + 9}', 'Sumber: Sri Harto, 1993: 168', normal_format)

    output.seek(0)  # Rewind the buffer
    return output

if __name__ == '__main__':
    app.run(debug=True)
