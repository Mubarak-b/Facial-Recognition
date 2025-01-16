from flask import Flask, render_template, request
import subprocess
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/run_main', methods=['POST'])
def run_main():
    try:
        subprocess.Popen(['python', 'Resources/temp.py'], cwd=os.path.dirname(os.path.realpath(__file__)))
        return 'Capturing Attendeance Please Wait'
    except Exception as e:
        return f'Error: {str(e)}'

if __name__ == '__main__':
    app.run(debug=True)
