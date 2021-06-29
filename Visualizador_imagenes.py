import pytesseract
from PIL import Image
import os
#Abrir un webbroser
import webbrowser
webbrowser.open('http://www.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaMovil.jsp')  # Go to example.com

import time
time.sleep(10)   # Delays for 5 seconds. You can also use a float value.

#Capturar pantalla activa
from PIL import ImageGrab
snapshot = ImageGrab.grab()
save_path = os.path.abspath(os.getcwd()) +r'\MySnapshot.jpg'
snapshot.save(save_path)

from PIL import Image

#Object detection
import cv2 as cv
import numpy as np
def MatchingMultipleObjects (allImage,findImage):
  img_rgb = cv.imread(allImage)
  img_gray = cv.cvtColor(img_rgb, cv.COLOR_BGR2GRAY)
  template = cv.imread(findImage,0)
  w, h = template.shape[::-1]
  res = cv.matchTemplate(img_gray,template,cv.TM_CCOEFF_NORMED)
  threshold = 0.8
  loc = np.where( res >= threshold) 
  #variable que cuenta el numero de objeto encontrado en la imagen
  numero_objeto=int(1)
  for pt in zip(*loc[::-1]):
    cv.rectangle(img_rgb, pt, (pt[0] + w, pt[1] + h), (0,0,255), 2)
    print('Ubicación del centro de la imagen #=%x' % numero_objeto + ' Centro X=%f' % (pt[0] + w/2) + ' Y=%d' % (pt[0] + h/2) )
    numero_objeto +=1
  cv.imwrite('resultado.jpg',img_rgb)
  print('Fin de Ejecución')

import os
#path current working directory
#ref:https://stackoverflow.com/questions/3430372/how-do-i-get-the-full-path-of-the-current-files-directory
#os.path.abspath(os.getcwd())
findImage_path= os.path.abspath(os.getcwd()) + r'\BuscarImagen\youtube.png'
MatchingMultipleObjects(save_path,findImage_path)
save_path_cropped=r'./resultado.jpg'




# Importamos Pytesseract
import pytesseract
codigo_captcha=pytesseract.image_to_string(save_path_cropped)
print(pytesseract.image_to_string(save_path_cropped))


#Click https://stackoverflow.com/questions/1181464/controlling-mouse-with-python
import ctypes

SetCursorPos = ctypes.windll.user32.SetCursorPos
mouse_event = ctypes.windll.user32.mouse_event

def left_click(x, y, clicks=1):
  SetCursorPos(x, y)
  for i in range(clicks):
   mouse_event(2, 0, 0, 0, 0)
   mouse_event(4, 0, 0, 0, 0)


#click texto Ruc
left_click(550, 275)
left_click(550, 275)
import win32com.client as comclt
wsh= comclt.Dispatch("WScript.Shell")
time.sleep(2) 


#wsh.SendKeys("111111111111")


time.sleep(2) 
#click texto captcha

left_click(930, 325) #left clicks at 200, 200 on your screen. Was able to send 10k clicks instantly
wsh.SendKeys(codigo_captcha[0:4])

#click boton buscar
left_click(820, 345)
