from selenium.webdriver.common.by import By
import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime
import pytz


url = 'https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094|26566160&rk=813013472&token=e6c72c6c-91c5-4247-ad08-454214e7f74e'

driver = webdriver.Chrome()
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

        driver.get('https://catalogo-vpfe.dian.gov.co/User/RedirectToBiller')

        elemento3 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'menu-button.documento'))
        )
        driver.execute_script("arguments[0].click();", elemento3)

        tiempo_espera_segundos = 10
        elemento = WebDriverWait(driver, tiempo_espera_segundos).until(
            EC.visibility_of_element_located((By.ID, 'OrderReference_ID'))
             )
        try:
            # -------------------------------------------------------- DATOS DEL EXCEL --------------------------------------------------------#
            excel_file_path = 'Datos.xlsx'
            workbook = openpyxl.load_workbook(excel_file_path)
            sheet = workbook.active
            data = []

            for row in sheet.iter_rows(min_row=2, values_only=True):
                data.append(row)
            print(data)

            for row in data:
                orden_compra = row[0]
                Tipo_cafe = row[1]
                doc_cl = row[2]
                nombre_cliente = (
                                     (str(row[5]) + " ") if (row[5] is not None) and str(row[5]) else "") + (
                                     (str(row[6]) + " ") if (row[6] is not None) and str(row[6]) else "") + (
                                     (str(row[3]) + " ") if (row[3] is not None) and str(row[3]) else "") + (
                                     (str(row[4]) + " ") if (row[4] is not None) and str(row[4]) else "")
                print(nombre_cliente)
                dir_cl = "VEREDA " + str(row[7])
                municipio = row[8]
                cod_postal = row[9]
                cant_kilos = row[10]
                valor = row[11]

                try:
                    id_del_select = 'OrderReference_ID'
                    nuevo_valor = orden_compra
                    script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                    driver.execute_script(script1)

                finally:
                    print("error")


                # -------------------------------------------------------- DATOS DEL DOCUMENTO --------------------------------------------------------#

                dropdown = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, 'RangoNum'))
                )
                # Crear un objeto Select para interactuar con el elemento de lista desplegable
                select = Select(dropdown)
                # Seleccionar una opción por valor
                select.select_by_value('a1622672-35d6-4132-a6a2-491657083a98')

                # id_del_select = 'OrderReference_ID'
                # # Nuevo valor que deseas establecer
                # nuevo_valor = '12345'
                # # Ejecutar JavaScript para cambiar el valor del elemento <select>
                # script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                # driver.execute_script(script1)

                elemento10 = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, 'ui-datepicker-trigger'))
                )
                elemento10.click()

                elemento20 = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, 'ui-state-default.ui-state-highlight.ui-state-hover'))
                )
                elemento20.click()

                id_del_select = 'paymentMeansField_PaymentMeansCode'
                # Nuevo valor que deseas establecer
                nuevo_valor = '1'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                id_del_select = 'TiopoNeg'
                # Nuevo valor que deseas establecer
                nuevo_valor = '1'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                # -------------------------------------------------------- DATOS DEL ADQUIRIENTE / COMPRADOR --------------------------------------------------------#

                elemento_datos = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseVendor"]'))
                )
                # Hacer clic en el elemento
                elemento_datos.click()

                id_del_select = 'accountingCustomerPartyField_Party_PartyTaxScheme_TaxLevelCode'
                # Nuevo valor que deseas establecer
                nuevo_valor = 'R-99-PN'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                id_del_select = 'ResponsabilidadTributariaAdquiriente'
                # Nuevo valor que deseas establecer
                nuevo_valor = 'ZZ'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                # -------------------------------------------------------- DATOS DEL VENDEDOR --------------------------------------------------------#

                elemento_vendedor = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseTwo"]'))
                )
                # Hacer clic en el elemento
                elemento_vendedor.click()

                codigo_postal = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.ID, 'accountingCustomerPartyField_Party_PhysicalLocation_Address_PostalZone'))
                )
                select_codigo_postal = Select(codigo_postal)
                select_codigo_postal.select_by_value('111711')

                id_del_select = 'origin'
                # Nuevo valor que deseas establecer
                nuevo_valor = '10'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script)

                id_del_select1 = 'RegimenFiscal'
                # Nuevo valor que deseas establecer
                nuevo_valor1 = 'R-99-PN'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
                driver.execute_script(script1)

                id_del_select1 = 'TipoDocumento'
                # Nuevo valor que deseas establecer
                nuevo_valor1 = '31'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
                driver.execute_script(script1)

                id_del_select1 = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_Country_IdentificationCode'
                # Nuevo valor que deseas establecer
                nuevo_valor1 = 'CO'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
                driver.execute_script(script1)

                id_del_select1 = 'NumeroDocumento'
                # Nuevo valor que deseas establecer
                nuevo_valor1 = doc_cl
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
                driver.execute_script(script1)

                id_del_select = 'Departamento'
                # Nuevo valor que deseas establecer
                nuevo_valor = '41'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>

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
                # Nuevo valor que deseas establecer
                nuevo_valor = '41551'

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
                # Nuevo valor que deseas establecer
                nuevo_valor = '2'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                id_del_select = 'ResposabilidadTributaria'
                # Nuevo valor que deseas establecer
                nuevo_valor = 'ZZ'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                id_del_select = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_AddressLine'
                # Nuevo valor que deseas establecer
                nuevo_valor = dir_cl
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                id_del_select = 'accountingCustomerPartyField_Party_PhysicalLocation_Address_PostalZone'
                # Nuevo valor que deseas establecer
                nuevo_valor = '417030'

                nuevo_texto = '417030'

                script = f"var select = document.getElementById('{id_del_select}'); " \
                         f"var option = document.createElement('option'); " \
                         f"option.value = '{nuevo_valor}'; " \
                         f"option.text = '{nuevo_texto}'; " \
                         f"select.add(option);"
                driver.execute_script(script)

                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                id_del_select = 'RazonSocial'
                # Nuevo valor que deseas establecer
                nuevo_valor = nombre_cliente
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                # -------------------------------------------------------- DETALLE DE PRODUCTO / SERVICIO --------------------------------------------------------#

                elemento_productos = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//a[@href="#collapseProductDetails"]'))
                )
                # Hacer clic en el elemento
                elemento_productos.click()

                id_del_select1 = 'Descripcion1'
                # Nuevo valor que deseas establecer
                nuevo_valor1 = Tipo_cafe
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
                driver.execute_script(script1)

                id_del_select1 = 'Codigo1'
                # Nuevo valor que deseas establecer
                nuevo_valor1 = '2'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
                driver.execute_script(script1)

                id_del_select1 = 'UM1'
                # Nuevo valor que deseas establecer
                nuevo_valor1 = 'KGM'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
                driver.execute_script(script1)

                id_del_select1 = 'PrecioUnitario1'
                # Nuevo valor que deseas establecer
                precio_unitario = '5'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{precio_unitario}';"
                driver.execute_script(script1)

                id_del_select1 = 'ImpuestosIVA1'
                # Nuevo valor que deseas establecer
                nuevo_valor1 = '0.00'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{nuevo_valor1}';"
                driver.execute_script(script1)

                id_del_select1 = 'Cantidad1'
                # Nuevo valor que deseas establecer
                cantidad = '1000'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select1}').value = '{cantidad}';"
                driver.execute_script(script1)

                id_del_select = 'Formagen1'
                # Nuevo valor que deseas establecer
                nuevo_valor = '1'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                id_del_select = 'FechaCompra1'
                # Nuevo valor que deseas establecer
                nuevo_valor = '28-12-2023'
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)

                zona_horaria = pytz.timezone('America/Bogota')

                # Obtener la fecha y hora actuales en la zona horaria deseada
                fecha_actual = datetime.now(zona_horaria)

                # Obtener la fecha local sin la información de la hora
                fecha_local = fecha_actual.date()

                id_del_select = 'ValorVentaItems1'
                # Nuevo valor que deseas establecer
                nuevo_valor = str(int(precio_unitario) * int(cantidad))
                # Ejecutar JavaScript para cambiar el valor del elemento <select>
                script1 = f"document.getElementById('{id_del_select}').value = '{nuevo_valor}';"
                driver.execute_script(script1)


        finally:
           print("error2")

finally:
    # Cerrar el navegador al finalizar
    input("Presiona Enter para salir...")
    driver.quit()
