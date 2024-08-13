from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Here you can process the form data
        form_data = request.form.to_dict()
        # Print or process the form data as needed
        print(form_data)
        # Redirect or render a confirmation template as needed
        return render_template('index.html', submitted=True)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
