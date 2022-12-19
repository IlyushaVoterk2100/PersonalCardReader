import pytesseract

def img2str(image):
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\TesseractORC\tesseract.exe'
    result = pytesseract.image_to_string(image)
    clean_result = ''
    for i in range(len(result)):
        if (result[i] == '0') or (result[i] == '1') or (result[i] == '2') or (result[i] == '3') or (result[i] == '4') or (result[i] == '5') or (result[i] == '6') or (result[i] == '7') or (result[i] == '8') or (result[i] == '9'):
            clean_result += str(result[i])
    if len(clean_result) == 10:
        return clean_result
