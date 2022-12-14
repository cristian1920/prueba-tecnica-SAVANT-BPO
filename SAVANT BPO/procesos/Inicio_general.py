from openpyxl import Workbook
import requests
import time
from email.mime.base import MIMEBase
from importlib.abc import FileLoader
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate, parseaddr 
from email import encoders
from selenium import webdriver
import chromedriver_autoinstaller
import  json
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

class InicioGeneral():

    def Conexion():
        should_start10=False
        while not should_start10:
            try:
                request = requests.get("http://www.google.com", timeout=5)
            except (requests.ConnectionError, requests.Timeout):
                print("Sin conexión a internet.")
            else:
                print("Con conexión a internet.")
                should_start10=True

    def EnvioCorreo(comunicado):
        # Iniciamos los parámetros del script
        remitente = 'camilo1211@hotmail.com'
        destinatario = ['cristian131411@gmail.com']
        # Creamos el objeto mensaje
        mensaje = MIMEMultipart()
        # Establecemos los atributos del mensaje
        mensaje['From'] = remitente
        mensaje['To'] = ", ".join(destinatario)
        mensaje['Subject'] = 'RPA Radicacion: Error'
        # Creamos el cuerpo del mensaje
        cuerpo = comunicado
        # Y lo agregamos al objeto mensaje como objeto MIME de tipo texto
        mensaje.attach(MIMEText(cuerpo, 'html'))
        # Creamos la conexión con el servidor
        sesion_smtp = smtplib.SMTP('smtp.office365.com', 587)
        # Ciframos la conexión
        sesion_smtp.starttls()
        # Iniciamos sesión en el servidor
        sesion_smtp.login('camilo1211@hotmail.com','Camilo1920*')
        # Convertimos el objeto mensaje a texto
        texto = mensaje.as_string()
        # Enviamos el mensaje
        InicioGeneral.Conexion()
        sesion_smtp.sendmail(remitente, destinatario, texto)
        # Cerramos la conexión
        sesion_smtp.quit()
        print('Mensaje enviado Exitosamente')
         
    #ingresa a amazon y valida los productos 
    def proceso1_():
        delay = 10
        InicioGeneral.Conexion()
        driver = webdriver.Chrome(ChromeDriverManager().install())       
        driver.maximize_window()
        time.sleep(4)
        InicioGeneral.Conexion()
        driver.get('https://www.amazon.com')
        InicioGeneral.Conexion()
        try:
            WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID ,'twotabsearchtextbox')))
            print ("Cargó input donde se ingresa el articulo marketplace amazon")
        except TimeoutException:
            print ("No lo cargo")
            mensaje = 'Falló en la carga del input donde se ingresa el articulo marketplace amazon'
            InicioGeneral.EnvioCorreo(mensaje)
        elem = driver.find_element(By.ID,'twotabsearchtextbox').send_keys('alexa echo dot 5th generation 2022',Keys.ENTER)
        InicioGeneral.Conexion()
        try:
            WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH ,'//*[@id="p_72/2661618011"]/span/a')))
            print ("Cargó el filtro de 4 estrellas o mas ")
        except TimeoutException:
            print ("No Cargó el filtro de 4 estrellas o mas")
            mensaje = 'Falló en la carga del select donde se ingresa nombre de usuario'
            InicioGeneral.EnvioCorreo(mensaje)
        elem = driver.find_element(By.XPATH,'//*[@id="p_72/2661618011"]/span/a').click()
        InicioGeneral.Conexion()
        time.sleep(5)
        try:
            WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID ,'low-price')))
            print ("Cargó el filtro de presio de 25 a 50 ")
            elem = driver.find_element(By.ID,'low-price').send_keys('25')
            elem = driver.find_element(By.ID,'high-price').send_keys('50')
            elem = driver.find_element(By.XPATH,'//*[@id="p_36/price-range"]/span/form/span[3]/span/input').click()
        except TimeoutException:
            print ("No Cargó el filtro de presio de 25 a 50")
            mensaje = 'fallo presionando sobre el link del filtro de presio de 25 a 50'
            InicioGeneral.EnvioCorreo(mensaje)
        
        InicioGeneral.Conexion()
        try:
                WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID ,'s-result-sort-select')))
                print ("Cargó el elemento, filtrar de menor a mayor precio")
        except TimeoutException:
            print ("No lo cargo")
            mensaje = 'Falló en la carga del filtrar de menor a mayor precio'
            InicioGeneral.EnvioCorreo(mensaje)
        select = Select(driver.find_element(By.ID,'s-result-sort-select'))
        select.select_by_value("price-asc-rank")
        calificacion=list()
        for i in range(1,4):
            try:
                WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH ,f'//*[@id="search"]/div[1]/div[1]/div/span[1]/div[1]/div[{i+1}]/div/div/div/div/div/div[2]/div/div/div[3]/div[1]/div/div[1]/div/a/span')))
                print ("Cargó el valor del producto ")
                valortext = driver.find_element(By.XPATH,f'//*[@id="search"]/div[1]/div[1]/div/span[1]/div[1]/div[{i+1}]/div/div/div/div/div/div[2]/div/div/div[3]/div[1]/div/div[1]/div/a/span').text
                print(valortext)
                valoraciontext = driver.find_element(By.XPATH,f'//*[@id="search"]/div[1]/div[1]/div/span[1]/div[1]/div[{i+1}]/div/div/div/div/div/div[2]/div/div/div[2]/div/span[1]').text
                print(valoraciontext)
            except TimeoutException:
                print ("No Cargó el el valor del producto")
                   
        
            try:
                WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH ,f'//*[@id="search"]/div[1]/div[1]/div/span[1]/div[1]/div[{i+1}]/div/div/div/div/div/div[2]/div/div/div[1]/h2/a/span')))
                print ("Cargó el nombre del producto ")
                nombretext = driver.find_element(By.XPATH,f'//*[@id="search"]/div[1]/div[1]/div/span[1]/div[1]/div[{i+1}]/div/div/div/div/div/div[2]/div/div/div[1]/h2/a/span').text
                print(nombretext)
            except TimeoutException:
                print ("No Cargó el nombre del producto")         
            verificar=valortext.replace("US$", "")
            if(float(verificar.replace("\n", "."))<=29):
                calificacion.append({
                                    'nombre': nombretext,
                                    'valoracion': valoraciontext,
                                    'precio':valortext.replace("\n", "."),
                                }) 
        with open('datos.json', 'w', encoding="utf-8") as file:
           json_s= json.dump(calificacion, file, indent=4, ensure_ascii=False)
        envio='Amazon'
        InicioGeneral.crearexcel(envio)
        InicioGeneral.proceso2(driver)
    #ingresa a aliexpress y valida los productos 
    def proceso2(driver):
        delay = 10
        time.sleep(4)
        InicioGeneral.Conexion()
        driver.get('https://es.aliexpress.com')
        InicioGeneral.Conexion()
        try:
            WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID ,'search-key')))
            print ("Cargó input donde se ingresa el articulo marketplace aliexpress")
        except TimeoutException:
            print ("No lo cargo")
            mensaje = 'Falló en la carga del input donde se ingresa el articulo marketplace aliexpress'
            InicioGeneral.EnvioCorreo(mensaje)
        elem = driver.find_element(By.ID,'search-key').send_keys('smart watch apple original series 7',Keys.ENTER)
        InicioGeneral.Conexion()
        try:
            WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH ,'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/span[2]/span[2]/label/span[1]/input')))
            print ("Cargó el filtro de 4 estrellas o mas ")
        except TimeoutException:
            print ("No Cargó el filtro de 4 estrellas o mas")
            mensaje = 'Falló en la carga del select donde se ingresa nombre de usuario'
            InicioGeneral.EnvioCorreo(mensaje)
        elem = driver.find_element(By.XPATH,'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/span[2]/span[2]/label/span[1]/input').click()
        InicioGeneral.Conexion()
        try:
            WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH ,'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/span[2]/span[2]/label/span[1]/input')))
            print ("Cargó tabla con los resultados ")
            calificacion=list()
            for i in range(1,4):
                try:
                    WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH ,f'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[2]/a[{i}]')))
                    print ("Cargó el producto ")
                    valortext = driver.find_element(By.XPATH,f'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[2]/a[{i}]/div[2]/div[1]').text
                    print(valortext)
                    try:
                        WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH ,f'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[2]/a[{i}]/div[2]/div[2]')))
                        try:
                            valoraciontext = driver.find_element(By.XPATH,f'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[2]/a[{i}]/div[2]/div[4]/span[2]').text
                            print(valoraciontext)
                            nombretext = driver.find_element(By.XPATH,f'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[2]/a[{i}]/div[2]/div[5]/h1').text
                            print(nombretext)
                        except:                                             
                            valoraciontext = driver.find_element(By.XPATH,f'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[2]/a[{i}]/div[2]/div[3]/span[2]').text
                            print(valoraciontext)
                            nombretext = driver.find_element(By.XPATH,f'//*[@id="root"]/div[1]/div/div[2]/div[2]/div/div[2]/a[{i}]/div[2]/div[4]/h1').text
                            print(nombretext)
                    except TimeoutException:
                        print ("no lo Cargó el producto ")
                    
                except TimeoutException:
                    print ("No Cargó el el valor del producto")
                    mensaje = 'Falló en la carga de los elementos del producto'
                    InicioGeneral.EnvioCorreo(mensaje)            
                verificar=valortext.replace("COP", "")
                print(str(verificar))
                if(str(verificar)<='50,000.00'):
                    calificacion.append({
                                        'nombre': nombretext,
                                        'valoracion': valoraciontext,
                                        'precio':valortext,
                                    }) 
            with open('datos.json', 'w', encoding="utf-8") as file:
                json_s= json.dump(calificacion, file, indent=4, ensure_ascii=False)
            envio='Aliexpress'
            InicioGeneral.crearexcel(envio)
            InicioGeneral.proceso3(driver)
        except TimeoutException:
            print ("No Cargó tabla con los resultados")
            mensaje = 'Falló en la carga de la tabla con los resultados'
            InicioGeneral.EnvioCorreo(mensaje)
        time.sleep(5)
    #ingresa a Ebay y valida los productos 
    def proceso3(driver):
        delay = 10
        time.sleep(4)
        InicioGeneral.Conexion()
        driver.get('https://www.ebay.com/sch/i.html?_from=R40&_nkw=parlantes+bluetooth&_sacat=0&_udhi=145.631&rt=nc&LH_AS=1')
        InicioGeneral.Conexion()
        time.sleep(5)
        calificacion=list()
        for i in range(1,3):
            try:
                WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME ,'s-item__price')))
                print ("Cargó el valor del producto ")
                valortext = driver.find_element(By.XPATH ,f'//*[@id="srp-river-results"]/ul/li[{i+1}]/div/div[2]/div[2]/div[1]/span').text
                print(valortext)
                valoraciontext = driver.find_element(By.XPATH,f'//*[@id="srp-river-results"]/ul/li[{i+1}]/div/div[2]/div[2]/span[1]/span/span[2]').text
                print(valoraciontext)
                nombretext = driver.find_element(By.XPATH,f'//*[@id="srp-river-results"]/ul/li[{i+1}]/div/div[2]/a/div/span').text
                print(nombretext)
            except TimeoutException:
                print ("No Cargó el el valor del producto")
                mensaje = 'Falló en la carga de los elementos del producto'
                InicioGeneral.EnvioCorreo(mensaje)    
        
                  
            verificar=valortext.replace("COP $", "")
            verificar=verificar.replace(" ", "")
            verificar=verificar.replace(".", "")
            if(int(verificar)<=20000000):
                calificacion.append({
                                    'nombre': nombretext,
                                    'valoracion': valoraciontext,
                                    'precio':valortext,
                                }) 
        with open('datos.json', 'w', encoding="utf-8") as file:
           json_s= json.dump(calificacion, file, indent=4, ensure_ascii=False)
        envio='Ebay'
        InicioGeneral.crearexcel(envio)
        InicioGeneral.proceso4(driver)
    #ingresa a Linio y valida los productos 
    def proceso4(driver):
        delay = 10
        time.sleep(4)
        InicioGeneral.Conexion()
        driver.get('https://www.linio.com.co/c/tv-y-video/televisores?price=9900-630000')
        InicioGeneral.Conexion()
        time.sleep(5)
        calificacion=list()
        for i in range(1,8):
            try:
                WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH ,'/html/body/div[3]/main/div[1]')))
                print ("Cargó busqueda del producto ")
                valortext = driver.find_element(By.XPATH ,f'//*[@id="catalogue-product-container"]/div[{i}]/a[1]/div[2]/div[3]/div[2]/div[1]/span').text
                print(valortext)
                try:
                    valoraciontext = driver.find_element(By.XPATH,f'//*[@id="catalogue-product-container"]/div[{i}]/a[2]/div/span[2]').text
                    print(valoraciontext)
                except:
                    valoraciontext=0
                nombretext = driver.find_element(By.XPATH,f'//*[@id="catalogue-product-container"]/div[{i}]/a[1]/div[2]/p/span').text
                print(nombretext)
            except TimeoutException:
                print ("No Cargó el el valor del producto")
                mensaje = 'Falló en la carga de los elementos del producto'
                InicioGeneral.EnvioCorreo(mensaje)    
        
                  
            verificar=valortext.replace("$", "")
            verificar=verificar.replace(".", "")
            if(int(verificar)<=500000):
                calificacion.append({
                                    'nombre': nombretext,
                                    'valoracion': valoraciontext,
                                    'precio':valortext,
                                }) 
        with open('datos.json', 'w', encoding="utf-8") as file:
           json_s= json.dump(calificacion, file, indent=4, ensure_ascii=False)
        envio='Linio'
        InicioGeneral.crearexcel(envio)
    #creamos el excel con los campos filtrados
    def crearexcel(envio):
        contador = 2
        time.sleep(1)
        book = Workbook()
        Productos = book.active
        Productos['A1']= "Nombre_Producto"
        Productos['B1'] = "Valoracion_Producto"
        Productos['C1'] = "Valor_Producto"
        Productos.title = "Productos"    
        with open('datos.json', encoding="utf-8") as file:  
            data = json.load(file)
            for a in data:
                Productos['A'+str(contador)] = str(a['nombre'])
                Productos['B'+str(contador)] = str(a['valoracion'])
                Productos['C'+str(contador)] = str(a['precio'])
                contador = contador+1    
            book.save(envio+'Productos.xlsx')
            InicioGeneral.correo(envio)
    #se envia correo con el archivo adjunto
    def correo(envio):   
        time.sleep(5) 
        remitente = 'camilo1211@hotmail.com' 
        destinatario = ['cristian131411@gmail.com','camilo1211@hotmail.com']
        # Creamos el objeto mensaje
        mensaje = MIMEMultipart()
        # Establecemos los atributos del mensaje
        mensaje['From'] = remitente
        mensaje['To'] = ", ".join(destinatario)
        mensaje['Subject'] = 'Prueba tecnica marketplace de '+envio
        fp=open(envio+'Productos.xlsx',"rb")
        adjunto= MIMEBase('multipart','encripte')
        adjunto.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(adjunto)
        adjunto.add_header('Content-Disposition','attachment',filename=envio+'Productos.xlsx')
        mensaje.attach(adjunto)
        # Creamos el cuerpo del mensaje
        cuerpo = 'Buen día <br>Envió Ecxel con los datos depurados de la pagina marketplace de '+envio+',<br> Adjunto la relación. <br><br>Quedo atento <br>Muchas gracias'
        # Y lo agregamos al objeto mensaje como objeto MIME de tipo texto
        mensaje.attach(MIMEText(cuerpo, 'html'))
        # Creamos la conexión con el servidor
        sesion_smtp = smtplib.SMTP('smtp.office365.com', 587)
        # Ciframos la conexión
        sesion_smtp.starttls()
        # Iniciamos sesión en el servidor
        sesion_smtp.login('camilo1211@hotmail.com','Camilo1920*')
        # Enviamos el mensaje
        sesion_smtp.sendmail(remitente, destinatario, str(mensaje))
        # Cerramos la conexión
        sesion_smtp.quit()
        print('Mensaje enviado Exitosamente')




        
      
        
