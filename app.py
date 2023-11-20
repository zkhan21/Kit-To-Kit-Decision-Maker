from flask import Flask, render_template, request
import partnumbercheck

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Retrieve the part numbers from the form
        part_number_a = request.form['part_number_a'].strip()
        part_number_b = request.form['part_number_b'].strip()

        # Remove dots and dashes
        part_number_a = part_number_a.replace('.', '').replace('-', '')
        part_number_b = part_number_b.replace('.', '').replace('-', '')

        # Process the cleaned part numbers
        results = partnumbercheck.get_part_number_results(part_number_a, part_number_b)
        return render_template('index.html', results=results, request=request)
    return render_template('index.html', results=None, request=request)

if __name__ == "__main__":
    app.run(debug=True)
