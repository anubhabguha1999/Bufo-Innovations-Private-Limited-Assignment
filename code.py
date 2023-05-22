from flask import Flask, request, jsonify
import xlwt

app = Flask(__name__)

@app.route('/createReport', methods=['POST'])
def create_report():
    try:
        request_data = request.get_json()
        filename = request_data.get('filename')
        data = request_data.get('data')
        if not filename or not data:
            return jsonify({'message': 'Invalid request payload'}), 400

        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('Sheet1')

        headers = list(data[0].keys())
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        for row, row_data in enumerate(data, start=1):
            for col, value in enumerate(row_data.values()):
                worksheet.write(row, col, value)

        workbook.save(filename + '.xls')

        return jsonify({'message': 'Excel file created successfully'}), 200

    except Exception as e:
        return jsonify({'message': str(e)}), 500

if __name__ == '__main__':
    app.run()
