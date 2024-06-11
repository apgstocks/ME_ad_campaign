import time
import pywhatkit
from docx import Document
import pyautogui
from pynput.keyboard import Key, Controller
import pandas as pd
from flask import Flask, render_template, request, jsonify,url_for
#from locator import *
keyboard = Controller()
app = Flask(__name__)

def getText(filename):
    doc = Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

# Function to send WhatsApp message
def send_whatsapp_message(msg, phone, img, start_index, end_index):
    try:
        print(f'start={start_index},end={end_index}')
        for ph in phone[start_index:end_index+1]:
            ph = '+91' + str(ph)
            print(f"Sending message to {ph}")
            try:
                success=pywhatkit.sendwhats_image(ph, img, msg, tab_close=True)
                print(success)
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
    return render_template('/index.html')

@app.route('/send-message', methods=['POST'])
def process_from():
    #image_path = image
    image_path = url_for('static', filename='ME_image.jpeg')
    #phone = ['7904223840']

    try:
        customer = pd.read_excel(excel, sheet_name="Sheet 1", engine='openpyxl')
        msg = getText(message)
        s = len(customer)
        
        
        new_index = pd.Series(range(1, s+1), name="serial_no")
        customer.set_index(new_index, inplace=True)
        print(customer.head())
        ustomer = customer.iloc[:, [0, 1]]
        #customer.set_index(new_index, inplace=True)
        #customer.drop(columns=['Unnamed: 2'], inplace=True)
        
        # Determine send option
        send_option = request.form.get('send-option')
        #print(request.form)
        #print(send_option)
        if send_option == 'selected':
            start_index = request.form.get('start_index')
            end_index = request.form.get('end_index')
            try:
                start_index = int(start_index)-1
                end_index = int(end_index)-1
            except ValueError as e:
                return jsonify({'error': f'Invalid indexes: {e}'}), 400
        else:
            start_index = 0
            end_index = s

        # Send WhatsApp message
        print('we are here in rooting for whatsapp messages')
        success = send_whatsapp_message(msg, customer['Phone_numbers'], image_path, start_index, end_index)
        if success:
            return jsonify({'message': 'Form submitted successfully'}), 200
        else:
            return jsonify({'error': 'Failed to send messages'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
