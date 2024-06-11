from flask import Flask, render_template, request, jsonify
import time
import pywhatkit
from docx import Document
import pyautogui
from pynput.keyboard import Key, Controller
import pandas as pd
import os

app = Flask(__name__)
keyboard = Controller()

def getText(filename):
    doc = Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

# Function to send WhatsApp message
def send_whatsapp_message(msg, phone, img, start_index, end_index):
    try:
        for ph in phone[start_index:end_index+1]:
            ph = '+91' + str(ph)
            print(f"Sending message to {ph}")
            try:
                pywhatkit.sendwhats_image(ph, img, msg, tab_close=True)
                time.sleep(2)
                pyautogui.click()
                keyboard.press(Key.enter)
                keyboard.release(Key.enter)
            except Exception as e:
                print(f"Message sending failed for {ph}: {e}")
        return True  # Indicate success
    except Exception as e:
        print(f"Error in sending messages: {e}")
        return False  # Indicate failure

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send-message', methods=['POST'])
def process_from():
    print('we are here')
    if 'docx_file' not in request.files or 'xlsx_file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    docx_file = request.files['docx_file']
    xlsx_file = request.files['xlsx_file']

    if docx_file.filename == '' or xlsx_file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    # Save files to a temporary location
    docx_path = os.path.join('uploads', docx_file.filename)
    xlsx_path = os.path.join('uploads', xlsx_file.filename)
    docx_file.save(docx_path)
    xlsx_file.save(xlsx_path)

    try:
        msg = getText(docx_path)
        customer = pd.read_excel(xlsx_path, sheet_name="Sheet 1", engine='openpyxl')
        s = len(customer)
        new_index = pd.Series(range(1, s+1), name="serial_no")
        customer.set_index(new_index, inplace=True)
        customer = customer.iloc[:, [0, 1]]


        # Determine send option
        send_option = request.form.get('send-option')
        if send_option == 'selected':
            start_index = request.form.get('start_index')
            end_index = request.form.get('end_index')
            try:
                start_index = int(start_index)
                end_index = int(end_index)
            except ValueError as e:
                return jsonify({'error': f'Invalid indexes: {e}'}), 400
        else:
            start_index = 1
            end_index = s

        # Path to the image in the static folder
        image_path = os.path.join(app.static_folder, 'ME_image.jpeg')  

        # Send WhatsApp message
        success = send_whatsapp_message(msg, customer['Phone_numbers'], image_path, start_index, end_index)
        if success:
            return jsonify({'message': 'Form submitted successfully'}), 200
        else:
            return jsonify({'error': 'Failed to send messages'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
