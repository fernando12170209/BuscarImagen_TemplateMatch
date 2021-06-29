
import time
import ctypes

import win32com.client as comclt
wsh= comclt.Dispatch("WScript.Shell")

SetCursorPos = ctypes.windll.user32.SetCursorPos
mouse_event = ctypes.windll.user32.mouse_event


def left_click(x, y, clicks=1):
  SetCursorPos(x, y)
  mouse_event(2, 0, 0, 0, 0)
  mouse_event(4, 0, 0, 0, 0)
   
def buscar(blq1,blq2,blq3,blq4):
    print("hola mundo")
    #primer bloque admite 11 digitos
    left_click(570, 200)
    left_click(570, 200)
    wsh.SendKeys(blq1)
    time.sleep(1) 
    #segundo bloque admite bloque 4 digitos
    left_click(570, 220)
    left_click(570, 220)
    wsh.SendKeys(blq2)
    time.sleep(1) 
    #tercer bloque admite 6 digitos
    left_click(570, 240)
    left_click(570, 240)
    wsh.SendKeys(blq3)
    time.sleep(1) 
    #cuarto bloque admite 4 digitos
    left_click(570, 280)
    left_click(570, 280)
    wsh.SendKeys(4)
    time.sleep(1) 
    
    #click en Fecha
    left_click(570, 300)
    left_click(570, 300)
    wsh.SendKeys(5)
    time.sleep(1) 
    
    #Nro de autorizacion
    left_click(570, 320)
    left_click(570, 320)
    wsh.SendKeys(6)
    time.sleep(1) 
    
    #click en consultar 
    left_click(750, 360)
    left_click(750, 360)
    wsh.SendKeys(6)
    time.sleep(1)
    
    #click en limpiar 
    left_click(850, 360)
    left_click(850, 360)
    wsh.SendKeys(6)
    time.sleep(1) 
    
    #definimos i como el numero de intentos para leer un captcha
    i=0
    while i <6:
        i+=1
        #Capturar pantalla activa
        from PIL import ImageGrab
        snapshot = ImageGrab.grab()
        save_path = "C:\\Users\\Fernando\\Desktop\\CURSOS\\Python\\ejercicios\\Fiscalia.jpg"
        snapshot.save(save_path)
        
        
        #https://pythontic.com/image-processing/pillow/crop
        # import the Python Image processing Library
        
        from PIL import Image
        
        # Create an Image object from an Image
        
        imageObject  = Image.open(save_path)
        
        # Crop the iceberg portion https://stackoverflow.com/questions/9983263/how-to-crop-an-image-using-pil
        w,h=imageObject.size
        cropped     = imageObject.crop((920,290,w-295,h-400))
        
        # Display the cropped portion
        
        #cropped.show()
        
        #
        save_path_cropped="C:\\Users\\Fernando\\Desktop\\CURSOS\\Python\\ejercicios\\Captcha_fiscalia.jpg"
        cropped.save(save_path_cropped)
        #from PIL import ImageOps
        #
        #border = (0, 30, 0, 30) # left, up, right, bottom
        #ImageOps.crop(imageObject, border)
        
        # Importamos Pytesseract
        import pytesseract
        codigo_captcha=pytesseract.image_to_string(save_path_cropped)
        print(pytesseract.image_to_string(save_path_cropped))
        
        
        
        
        if len(codigo_captcha)==5 and codigo_captcha.isdigit()==True:
                print("codigo captcha correcto")
                i=7
                #salir del bucle por que ya se puede acceder
                
                #click en texbox para escribir el captcha
                left_click(920, 300)
                
                time.sleep(1) 
                #escribir captcha
                wsh.SendKeys(codigo_captcha[:5])
                
                
                #click boton consultar
                left_click(1080, 300)
                time.sleep(2)
                # si: El codigo captcha es invalido!!
                snapshot = ImageGrab.grab()
                save_path = "C:\\Users\\Fernando\\Desktop\\CURSOS\\Python\\ejercicios\\Fiscalia_resultado.jpg"
                snapshot.save(save_path)
                
                if "El codigo captcha es invalido!!"  in pytesseract.image_to_string(save_path):
                    #seguir buscando hasta escribir un captcha correcto
                    print("captcha equivocado, accrion:rehacer")
                    left_click(850, 150)
                    i=1
                #break
        else:
                print("codigo captcha incorrecto")
                if codigo_captcha[:5].isdigit()==True:
                    i=7
                    print("codigo captcha correcto")
                    #salir del bucle por que ya se puede acceder
                    #click en texbox para escribir el captcha
                    left_click(920, 300)
                    
                    time.sleep(1) 
                    #escribir captcha
                    wsh.SendKeys(codigo_captcha[:5])
                    
                    
                    #click boton consultar
                    left_click(1080, 300)
                    time.sleep(2)
                    # si: El codigo captcha es invalido!!
                    snapshot = ImageGrab.grab()
                    save_path = "C:\\Users\\Fernando\\Desktop\\CURSOS\\Python\\ejercicios\\Fiscalia_resultado.jpg"
                    snapshot.save(save_path)
                    
                    if "El codigo captcha es invalido!!"  in pytesseract.image_to_string(save_path):
                        #seguir buscando hasta escribir un captcha correcto
                        print("captcha equivocado, accrion:rehacer")
                        left_click(850, 150)
                        i=1
                    #break
                else:
                    left_click(1020, 300)
                    left_click(1020, 300)
                    time.sleep(2) 
                    #refresh captcha
        
    #click en texbox para escribir el captcha
    #left_click(920, 300)
    
    time.sleep(1) 
    #escribir captcha
    #wsh.SendKeys(codigo_captcha[:5])
    
    
    #click boton consultar
    #left_click(1080, 300)
    # si: El codigo captcha es invalido!!
    #snapshot = ImageGrab.grab()
    #save_path = "C:\\Users\\Fernando\\Desktop\\CURSOS\\Python\\ejercicios\\Fiscalia_resultado.jpg"
    #snapshot.save(save_path)
    #print(pytesseract.image_to_string(save_path))
    
    if "No existen registros con los datos ingresados"  in pytesseract.image_to_string(save_path):
            print("El numero de caso no es valido")
    else:
        print("Ahora le damos click en ver")

blq1=1
blq2=2018
blq3=2019
blq4=1
while blq4<=3:
    buscar(blq1,blq2,blq3,blq4)
    blq4+=1