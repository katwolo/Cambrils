import os
import openai
import requests
from flask import Flask, request

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
openai.api_key = OPENAI_API_KEY

app = Flask(__name__)

def start_message():
    return "👋 Bon dia! Sóc Pixelet, el teu assistent tecnològic valencià. En què et puc ajudar avui?"

def send_message(chat_id, text):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    requests.post(url, json={"chat_id": chat_id, "text": text})

def get_pixelet_response(user_input):
    prompt = f"""
Eres Pixelet, un assistent valencià, expert en tecnologia i educació digital. Parles amb entusiasme, en valencià, ets proper i simpàtic. Sempre acabes amb una frase inspiradora com “Pedra a pedra, farem paret.”

Usuari: {user_input}
Pixelet:"""
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        max_tokens=250,
        temperature=0.7,
    )
    return response.choices[0].text.strip()

@app.route(f"/{TELEGRAM_TOKEN}", methods=["POST"])
def webhook():
    data = request.get_json()
    if "message" in data and "text" in data["message"]:
        chat_id = data["message"]["chat"]["id"]
        text = data["message"]["text"]
        if text == "/start":
            send_message(chat_id, start_message())
        else:
            reply = get_pixelet_response(text)
            send_message(chat_id, reply)
    return {"ok": True}

@app.route("/", methods=["GET"])
def index():
    return "Pixelet està en marxa!"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
