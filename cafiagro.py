from selenium.webdriver.common.by import By
import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

# Ruta del archivo Excel
excel_file_path = 'Datos.xlsx'

# URL del sitio web
url = 'https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094|26566160&rk=813013472&token=307268dd-a4c5-4660-a232-4f6cde063687'


# Función para leer datos de Excel
def leer_datos_desde_excel():
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active
    data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(row)
    return data

def ingresa_data(data):
    # Configurar Selenium
    driver = webdriver.Chrome()
    driver.get('https://www.facebook.com/')
    # Utilizar los datos en Selenium
    for row in data:
        # Aquí puedes realizar interacciones con Selenium usando los datos
        # Por ejemplo, enviar datos a un formulario, hacer clic en botones, etc.
        input_element = driver.find_element_by_id('email')
        input_element.send_keys(row[0])

    # Realizar más acciones con Selenium según sea necesario

driver = webdriver.Chrome()
#try:
    #datos_excel = leer_datos_desde_excel(excel_file_path)
    #ingresa_data(datos_excel)
#finally:
    # Cerrar el navegador al finalizar
    #input(":(((")
# Configurar Selenium y abrir el sitio web


def iniciar_sesion():
    elemento = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, 'legalRepresentative'))
    )
    driver.execute_script("arguments[0].click();", elemento)

    NIT_Representante_Legal = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'UserCode'))
    )
    NIT_Representante_Legal.send_keys('26566160')

    NIT_Empresa = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'CompanyCode'))
    )
    NIT_Empresa.send_keys('813013472')

    elemento1 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, 'btn-primary'))
    )
    driver.execute_script("arguments[0].click();", elemento1)


try:

    driver.get(url)

    try:
        iniciar_sesion()

    finally:

        # elemento1 = WebDriverWait(driver, 10).until(
        #    EC.element_to_be_clickable((By.ID, 'mainnav-toggle'))
        # )
        # driver.execute_script("arguments[0].click();", elemento1)
        
        # elemento2 = WebDriverWait(driver, 10).until(
        #    EC.element_to_be_clickable((By.ID, 'Invoice'))
        # )
        # driver.execute_script("arguments[0].click();", elemento2)

        # elemento4 = WebDriverWait(driver, 10).until(
        #    EC.element_to_be_clickable((By.ID, 'Users'))
        # )
        # driver.execute_script("arguments[0].click();", elemento4)


        driver.get('https://catalogo-vpfe.dian.gov.co/User/RedirectToBiller')


        elemento3 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'menu-button.documento'))
        )
        driver.execute_script("arguments[0].click();", elemento3)

        

        elemento_vendedor = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseTwo"]'))
        )
        # Hacer clic en el elemento
        elemento_vendedor.click()

        codigo_postal = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'accountingCustomerPartyField_Party_PhysicalLocation_Address_PostalZone'))
        )
        select_codigo_postal = Select(codigo_postal)
        # Seleccionar una opción por valor
        select_codigo_postal.select_by_value('111711')

        id_del_select = 'origin'
        # Nuevo valor que deseas establecer
        nuevo_valor = '10'
        # Ejecutar JavaScript para cambiar el valor del elemento <select>
        script = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script)

        id_del_select1 = 'RegimenFiscal'
         #Nuevo valor que deseas establecer
        nuevo_valor1 = 'R-99-PN'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
        driver.execute_script(script1)

        id_del_select1 = 'TipoDocumento'
         #Nuevo valor que deseas establecer
        nuevo_valor1 = '31'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
        driver.execute_script(script1)

        id_del_select1 = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_Country_IdentificationCode'
         #Nuevo valor que deseas establecer
        nuevo_valor1 = 'CO'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
        driver.execute_script(script1)

        id_del_select1 = 'NumeroDocumento'
         #Nuevo valor que deseas establecer
        nuevo_valor1 = '12143849'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
        driver.execute_script(script1)

        id_del_select = 'Departamento'
         #Nuevo valor que deseas establecer
        nuevo_valor = '41'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        
        # Nuevo valor y texto para el elemento <option>

        nuevo_texto = '41 | Huila'

        # Ejecutar JavaScript para insertar el nuevo elemento <option>
        script = f"var select = document.getElementById('{id_del_select}'); " \
                f"var option = document.createElement('option'); " \
                f"option.value = '{nuevo_valor}'; " \
                f"option.text = '{nuevo_texto}'; " \
                f"select.add(option);"
        driver.execute_script(script)

        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        

        id_del_select = 'Ciudad'
         #Nuevo valor que deseas establecer
        nuevo_valor = '41551'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        
        # Nuevo valor y texto para el elemento <option>

        nuevo_texto = '41551 | PITALITO'

        # Ejecutar JavaScript para insertar el nuevo elemento <option>
        script = f"var select = document.getElementById('{id_del_select}'); " \
                f"var option = document.createElement('option'); " \
                f"option.value = '{nuevo_valor}'; " \
                f"option.text = '{nuevo_texto}'; " \
                f"select.add(option);"
        driver.execute_script(script)

        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)


        id_del_select = 'tipoContribuyente'
         #Nuevo valor que deseas establecer
        nuevo_valor = '2'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        id_del_select = 'ResposabilidadTributaria'
         #Nuevo valor que deseas establecer
        nuevo_valor = 'ZZ'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        id_del_select = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_AddressLine'
         #Nuevo valor que deseas establecer
        nuevo_valor = 'Verede Quinchana'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        id_del_select = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_PostalZone'
         #Nuevo valor que deseas establecer
        nuevo_valor = '417030'

        nuevo_texto = '417030'


        script = f"var select = document.getElementById('{id_del_select}'); " \
                f"var option = document.createElement('option'); " \
                f"option.value = '{nuevo_valor}'; " \
                f"option.text = '{nuevo_texto}'; " \
                f"select.add(option);"
        driver.execute_script(script)

         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)


        id_del_select = 'RazonSocial'
         #Nuevo valor que deseas establecer
        nuevo_valor = 'Jerson Sanchez'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        elemento_vendedor = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseTwo"]'))
        )
        # Hacer clic en el elemento
        elemento_vendedor.click()


        #-------------------------------------------------------- DATOS DEL DOCUMENTO --------------------------------------------------------#
        
        elemento_vendedor = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseOne"]'))
        )
        # Hacer clic en el elemento
        elemento_vendedor.click()

        dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'RangoNum'))
        )
        # Crear un objeto Select para interactuar con el elemento de lista desplegable
        select = Select(dropdown)
        # Seleccionar una opción por valor
        select.select_by_value('a1622672-35d6-4132-a6a2-491657083a98')


        id_del_select = 'OrderReference_ID'
         #Nuevo valor que deseas establecer
        nuevo_valor = '12345'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        # id_del_input_fecha = 'OrderReferenceIssueDate'

        # # Valor de fecha que deseas establecer
        # valor_fecha = '2023-12-31'

        # # Encontrar el campo de entrada de tipo fecha y enviar el valor
        # input_fecha = driver.find_element_by_id(id_del_input_fecha)
        # input_fecha.clear()  # Limpiar cualquier valor existente
        # input_fecha.send_keys(valor_fecha)

        elemento10 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, 'ui-datepicker-trigger'))
        )
        elemento10.click()

        elemento20 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, 'ui-state-default.ui-state-highlight.ui-state-hover'))
        )
        elemento20.click()

        id_del_select = 'paymentMeansField_PaymentMeansCode'
         #Nuevo valor que deseas establecer
        nuevo_valor = '1'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        id_del_select = 'TiopoNeg'
         #Nuevo valor que deseas establecer
        nuevo_valor = '1'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        elemento_vendedor = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseOne"]'))
        )
        # Hacer clic en el elemento
        elemento_vendedor.click()

        #-------------------------------------------------------- DATOS DEL ADQUIRIENTE / COMPRADOR --------------------------------------------------------#

        elemento_vendedor = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseVendor"]'))
        )
        # Hacer clic en el elemento
        elemento_vendedor.click()

        id_del_select = 'accountingCustomerPartyField_Party_PartyTaxScheme_TaxLevelCode'
         #Nuevo valor que deseas establecer
        nuevo_valor = 'R-99-PN'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        id_del_select = 'ResponsabilidadTributariaAdquiriente'
         #Nuevo valor que deseas establecer
        nuevo_valor = 'ZZ'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)


        input_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'Codigo1'))
        )

        input_element.send_keys('2')

        input_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'Descripcion1'))
        )

        input_element.send_keys('CAFE SECO')

        input_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'UM1'))
        )

        input_element.send_keys('KGM')


        input_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'PrecioUnitario1'))
        )

        input_element.send_keys('5')

        

        input_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ImpuestosIVA1'))
        )

        input_element.send_keys('0,00')

        input_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'Cantidad1'))
        )

        input_element.send_keys('1000')

        id_del_select = 'Formagen1'
         #Nuevo valor que deseas establecer
        nuevo_valor = '1'
         #Ejecutar JavaScript para cambiar el valor del elemento <select>
        script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
        driver.execute_script(script1)

        

        
        
        



        
        


        
        

        

        





        







finally:
    # Cerrar el navegador al finalizar
    input("Presiona Enter para salir...")
    driver.quit()
