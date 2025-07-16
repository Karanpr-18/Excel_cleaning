from app import app

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

# To run this Flask application, you can use the following commands:
# python main.py
# or
# venv\Scripts\activate
# pip install -r requirements.txt
# gunicorn -w 4 -b 0.0.0.0:5000 app:app