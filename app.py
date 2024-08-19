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

    # Creating DataFrames based on the form data
    # Example for Station Info
    station_info = {
        'Field': ['NAMA STASIUN', 'Wilayah Sungai', 'Kode Database', 'Kode Stasiun', 
                  'Kelurahan/Desa', 'Tahun Pendirian', 'Lintang Selatan', 'Kecamatan',
                  'Tipe Alat', 'Bujur Timur', 'Kab/Kota', 'Pengelola', 'Elevasi'],
        'Value': [
            data.get('nama_stasiun', '[Nama Stasiun]'),
            data.get('wilayah_sungai', '[Wilayah Sungai]'),
            data.get('kode_database', '[Kode Database]'),
            data.get('kode_stasiun', '[Kode Stasiun]'),
            data.get('kelurahan', '[Kelurahan/Desa]'),
            data.get('tahun_pendirian', '[Tahun Pendirian]'),
            data.get('longitude', '[Longitude]'),
            data.get('kecamatan', '[Kecamatan]'),
            data.get('tipe_alat', '[Tipe Alat]'),
            data.get('latitude', '[Latitude]'),
            data.get('kabupaten', '[Kabupaten]'),
            data.get('pengelola', '[Pengelola]'),
            data.get('elevasi', '[Elevation]')
        ]
    }
    df_station_info = pd.DataFrame(station_info)

    # Example for Rainfall Data
    days = list(range(1, 32))
    months = list(range(1, 13))
    df_rainfall_data = pd.DataFrame({'Tanggal': days})
    for month in months:
        df_rainfall_data[f'Bulan {month}'] = [data.get(f'day{day}_month{month}', 0) for day in days]

    # Example for Totals
    df_totals = pd.DataFrame({
        'Metric': ['Total', 'Periode1', 'Periode2', 'Periode3', 'Maksimum'],
        **{f'Bulan {month}': [
            data.get(f'total_month{month}', 0),
            data.get(f'periode1_month{month}', 0),
            data.get(f'periode2_month{month}', 0),
            data.get(f'periode3_month{month}', 0),
            data.get(f'maksimum_month{month}', 0),
        ] for month in months}
    })

    # Example for Data Hujan
    df_data_hujan = pd.DataFrame({
        'Data Hujan': [f'Bulan {month}' for month in months],
        'Value': [data.get(f'datahujan_month{month}', 0) for month in months]
    })

    # Example for Analysis
    analysis_columns = ['No', 'Bulan', 'Curah Hujan', 'Sk*', '[Sk*]', 'Dy^2', 'Dy', 'Sk**', '[Sk**]']
    analysis_data = [
        [
            i + 1,
            month,
            data.get(f'curah_hujan_{i}', 0),
            data.get(f'sk_{i}', 0),
            data.get(f'sk_brackets_{i}', 0),
            data.get(f'dy2_{i}', 0),
            data.get(f'dy_{i}', 0),
            data.get(f'sk_star_{i}', 0),
            data.get(f'sk_star_brackets_{i}', 0)
        ]
        for i, month in enumerate(['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'])
    ]
    df_analysis = pd.DataFrame(analysis_data, columns=analysis_columns)

    # Writing to an Excel file
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_station_info.to_excel(writer, index=False, sheet_name='Station Info')
        df_rainfall_data.to_excel(writer, index=False, sheet_name='Rainfall Data')
        df_totals.to_excel(writer, index=False, sheet_name='Totals')
        df_data_hujan.to_excel(writer, index=False, sheet_name='Data Hujan')
        df_analysis.to_excel(writer, index=False, sheet_name='Analysis')

        # Example of formatting: Setting column widths
        workbook = writer.book
        worksheet = writer.sheets['Station Info']
        worksheet.set_column('A:B', 20)
        worksheet = writer.sheets['Rainfall Data']
        worksheet.set_column('A:N', 15)

    output.seek(0)  # Rewind the buffer
    return output

if __name__ == '__main__':
    app.run(debug=True)
