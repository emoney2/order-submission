# submit_server.py
from flask import Flask, render_template_string
app = Flask(__name__)

@app.route("/")
def index():
    # This is where you could import and call into your real order-submission logic,
    # but for now let's just proxy to the placeholder or render a custom template.
    return render_template_string("""
      <h1>Order Submission App</h1>
      <p>This is still a placeholder – we’ll wire up your real code next.</p>
    """)