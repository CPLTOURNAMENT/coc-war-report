from flask import Flask, send_from_directory

app = Flask(__name__, static_folder='static')

@app.route('/')
def index():
    return "Clash of Clans War Data Server Running"

@app.route('/live_war_auto_update.xlsx')
def serve_excel():
    return send_from_directory('static', 'live_war_auto_update.xlsx')

app.run(host='0.0.0.0', port=81)
