from PIL import Image
import os
import sys
#Object detection
import cv2 as cv
import numpy as np

#Click windows
import ctypes
#https://dotnetcoretutorials.com/2018/07/21/uploading-images-in-a-pure-json-api/

#OCR
#handwriting tensorflow
#https://www.pyimagesearch.com/2020/08/24/ocr-handwriting-recognition-with-opencv-keras-and-tensorflow/
def ResizeImage(findImage_path):
    src = cv.imread(findImage_path, cv.IMREAD_UNCHANGED)
    #Reajusto el tamaño de la imagen (desde el 10% al 120% de tamaño de la imagen)
    #linspace interpola dese 0.1 hasta 1.2 en 15 pasos
    #[::-1] se invierte el rango (empieza desde el ultimo registro y termina al inicio)
    for scale in np.linspace(0.4, 0.8, 15)[::-1]:
        width = int(src.shape[1] * scale)
        height = int(src.shape[0] * scale)
        # dsize
        dsize = (width, height)
        output = cv.resize(src, dsize)
        #cv.imshow("imagen",output)
        #cv.waitKey(0)

def MatchingMultipleObjects (allImage,findImage):
    img_rgb = cv.imread(allImage)    
    img_gray = cv.cvtColor(img_rgb, cv.COLOR_BGR2GRAY)
    #cv.imshow('image',img_rgb)
    #cv.waitKey(0)

    #template = cv.imread(findImage,0)
    template_rgb = cv.imread(findImage)
    template = cv.cvtColor(template_rgb, cv.COLOR_BGR2GRAY)

    print(img_gray.shape[::-1])
    w, h = template.shape[::-1]
    wi,he=img_gray.shape[::-1]
    #Reajusto el tamaño de la imagen (desde el 10% al 130% de tamaño de la imagen)
    #linspace interpola dese 0.2 hasta 1.3 en 20 pasos
    #[::-1] se invierte el rango (empieza desde el ultimo registro y termina al inicio)
    for scale in np.linspace(0.2, 1.3, 20)[::-1]:
        width = int(template.shape[1] * scale)
        height = int(template.shape[0] * scale)
        
        if(width<wi and height<he):
            w, h = template.shape[::-1]
            dsize=(int(w*scale),int(h*scale))
            #dsize = (width, height)
            output = cv.resize(template, dsize)
            w, h = output.shape[::-1]

            #cv.imshow('image',output)
            #cv.waitKey(0)
            #output_read=cv.imread(output,0)
            res = cv.matchTemplate(img_gray,output,cv.TM_CCOEFF_NORMED)
            threshold = 0.8
            loc = np.where( res >= threshold) 
            #variable que cuenta el numero de objeto encontrado en la imagen
            numero_objeto=int(1)
            for pt in zip(*loc):
                # Start coordinate, here (100, 50)
                # represents the top left corner of rectangle
                start_point = (pt[1],pt[0])   
                print('start_point:' +str(start_point))             
                # Ending coordinate, here (125, 80)
                # represents the bottom right corner of rectangle
                end_point = (pt[1] + w, pt[0] + h)
                print('end_point:' +str(end_point))  
                cv.rectangle(img_rgb, start_point, end_point, (0,0,255), 2)
                cv.imwrite('resultado'+str(numero_objeto)+'.jpg',img_rgb)
                #Image center coordinates
                print('Ubicación del centro de la imagen #=%x' % numero_objeto + ' Centro X=%f' % (pt[1] + w/2) + ' Y=%d' % (pt[0] + h/2) )
                numero_objeto +=1

    """

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
    
    """

def pathExit(pathUser):
    n=os.path.exists(pathUser)
    return n





def left_click(x, y, clicks=1):
    SetCursorPos = ctypes.windll.user32.SetCursorPos
    mouse_event = ctypes.windll.user32.mouse_event
    SetCursorPos(x, y)
    for i in range(clicks):
        mouse_event(2, 0, 0, 0, 0)
        mouse_event(4, 0, 0, 0, 0)
if __name__ == "__main__":
    #
    imageScreen= os.path.abspath(os.getcwd()) + r'\BuscarImagen\resultado.jpg'
    if pathExit(imageScreen)==False:
        sys.exit()
    findImage=os.path.abspath(os.getcwd()) + r'\BuscarImagen\Refresh_ConsultaRucMovil2.jpg'
    if pathExit(findImage)==False:
        sys.exit()
    MatchingMultipleObjects(imageScreen,findImage)

    #left_click(550, 275)
    #left_click(275, 275)

