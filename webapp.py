from flask import Flask

app = Flask(__name__)

@app.route("/")
def index():
    return "<h1>Order Submission App</h1><p>This is a placeholder.</p>"

if __name__ == "__main__":
    # listen on all interfaces so Render can reach it
    app.run(host="0.0.0.0", port=5000)
