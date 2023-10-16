from flask import Flask, request, jsonify
from library import HTMLComponentsExtractor,extract_html_and_css  # Replace with your library name

app = Flask(__name__)

@app.route('/convert-html-to-docx', methods=['POST'])
def convert_html_to_docx():
    data = request.get_json()
    html_content = data.get('htmlContent')
    if html_content:
        docx_content = HTMLComponentsExtractor(html_content)  # Replace with the actual conversion function
        return jsonify({'docxContent': docx_content})
    else:
        return jsonify({'error': 'Invalid request'}), 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)  # Adjust host and port as needed
