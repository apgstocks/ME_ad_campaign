import os
import platform

if 'DISPLAY' not in os.environ and platform.system() != 'Windows':
    os.environ['DISPLAY'] = ':0'  # Mock DISPLAY for headless environment

from flask import Flask, render_template, request, jsonify
import time
import pywhatkit
from docx import Document
import pyautogui
from pynput.keyboard import Key, Controller
import pandas as pd
import logging

app = Flask(__name__)
keyboard = Controller()

# Set up logging
logging.basicConfig(level=logging.INFO)

def getText(filename):
    try:
        doc = Document(filename)
        fullText = []
        for para in doc.paragraphs:
            fullText.append(para.text)
        return '\n'.join(fullText)
    except Exception as e:
        logging.error(f"Error reading DOCX file: {e}")
        raise

def send_whatsapp_message(msg, phone, img, start_index, end_index):
    try:
        for ph in phone[start_index:end_index+1]:
            ph = '+91' + str(ph)
            logging.info(f"Sending message to {ph}")
            try:
                pywhatkit.sendwhats_image(ph, img, msg, tab_close=True)
                time.sleep(2)
                pyautogui.click()
                keyboard.press(Key.enter)
                keyboard.release(Key.enter)
            except Exception as e:
                logging.error(f"Message sending failed for {ph}: {e}")
        return True
    except Exception as e:
        logging.error(f"Error in sending messages: {e}")
        return False

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send-message', methods=['POST'])
def process_form():
    if 'docx_file' not in request.files or 'xlsx_file' not in request.files:
        logging.error("No file part")
        return jsonify({'error': 'No file part'}), 400

    docx_file = request.files['docx_file']
    xlsx_file = request.files['xlsx_file']

    if docx_file.filename == '' or xlsx_file.filename == '':
        logging.error("No selected file")
        return jsonify({'error': 'No selected file'}), 400

    try:
        # Save files to a temporary location
        docx_path = os.path.join('uploads', docx_file.filename)
        xlsx_path = os.path.join('uploads', xlsx_file.filename)
        os.makedirs('uploads', exist_ok=True)
        docx_file.save(docx_path)
        xlsx_file.save(xlsx_path)

        msg = getText(docx_path)
        customer = pd.read_excel(xlsx_path, sheet_name="Sheet 1", engine='openpyxl')
        s = len(customer)
        new_index = pd.Series(range(1, s+1), name="serial_no")
        customer.set_index(new_index, inplace=True)
        customer.drop(columns=['Unnamed: 2'], inplace=True)

        # Determine send option
        send_option = request.form.get('send-option')
        if send_option == 'selected':
            start_index = request.form.get('start_index')
            end_index = request.form.get('end_index')
            try:
                start_index = int(start_index)
                end_index = int(end_index)
            except ValueError as e:
                logging.error(f"Invalid indexes: {e}")
                return jsonify({'error': f'Invalid indexes: {e}'}), 400
        else:
            start_index = 1
            end_index = s

        # Path to the image in the static folder
        image_path = os.path.join(app.static_folder, 'image.jpg')

        # Send WhatsApp message
        phone_numbers = customer['Phone_numbers'].tolist()
        success = send_whatsapp_message(msg, phone_numbers, image_path, start_index, end_index)
        if success:
            return jsonify({'message': 'Form submitted successfully'}), 200
        else:
            logging.error("Failed to send messages")
            return jsonify({'error': 'Failed to send messages'}), 500

    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        return jsonify({'error': str(e)}), 500
    finally:
        # Clean up uploaded files
        if os.path.exists(docx_path):
            os.remove(docx_path)
        if os.path.exists(xlsx_path):
            os.remove(xlsx_path)

if __name__ == '__main__':
    app.run(debug=True)
