from playwright.sync_api import Playwright, sync_playwright, expect
from datetime import datetime
import json
import re
import time

def funcion_registro_clasificacion(playwright: Playwright,URL_PANDORA,USERNAME,PASSWORD,curso_codigo) -> None:
    
    print("Iniciando registro de clasificación...")
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    page.goto(URL_PANDORA)
    page.locator("#username").fill(USERNAME)
    page.locator("#password").fill(PASSWORD)
    #page.locator("#password").press("Enter")
    page.get_by_role("link").filter(has_text=re.compile(r"^$")).press("Enter")
    
    page.goto("https://ares.pucp.edu.pe/pucp/jsp/Intranet.jsp?url=%2Fpucp%2Fidiomas%2Fidwpanal%2FidwpanalPucpQuestionaccion%3DVerPanel")
    page.goto("https://ares.pucp.edu.pe/pucp/jsp/Intranet.jsp?url=%2Fpucp%2Fidiomas%2Fidwclasf%2FidwclasfPucpQuestionaccion%3DMostrarClasificacion")
    resultados_registro = []
    longitud = len(curso_codigo)
    for i in range(0, longitud):
        try:
            print(f"Registrando clasificación para el alumno {curso_codigo[i][1]} y curso {curso_codigo[i][0]}...")
            page.locator("iframe[name=\"frame_mid\"]").content_frame.locator("#codAlumno").fill(curso_codigo[i][1])
            page.locator("iframe[name=\"frame_mid\"]").content_frame.locator("#codAlumno").press("Tab")
            print("Habilitando edición del registro...")
            time.sleep(2)  # Esperar a que se cargue la información del alumno
            page.locator("iframe[name=\"frame_mid\"]").content_frame.locator("div").filter(has_text="Grabar Buscar Editar Regresar").locator("button[name=\"btnEditar\"]").click()    
            # obtener todas las opciones del combo '#cmbCurso' dentro del iframe 'frame_mid'
            print("Obteniendo opciones de cursos...")
            frame = page.frame(name="frame_mid")
            options = frame.locator("#cmbCurso")
            options = options.evaluate("sel => Array.from(sel.options).map(o => ({ value: o.value, text: o.text.trim() }))")
            for opt in options:
                if curso_codigo[i][0] in opt["text"]:
                    codigo_curso = str(opt["value"])
                    print(f"Codigo de curso encontrado: {opt['value']} para el curso {opt['text']}")
                    break
            else:
                print(f"No se encontró {curso_codigo[i][0]} en las opciones de curso.")

            print("Seleccionando curso, sede y registrando observación...")
            page.locator("iframe[name=\"frame_mid\"]").content_frame.locator("#cmbCurso").select_option(codigo_curso)
            page.locator("iframe[name=\"frame_mid\"]").content_frame.locator("#cmbLocal").select_option("5")
            page.locator("iframe[name=\"frame_mid\"]").content_frame.locator("textarea[name=\"observaciones\"]").click()
            page.locator("iframe[name=\"frame_mid\"]").content_frame.locator("textarea[name=\"observaciones\"]").fill(curso_codigo[i][2])
            page.once("dialog", lambda dialog: dialog.accept())
            print("Guardando registro...")
            page.locator("iframe[name=\"frame_mid\"]").content_frame.get_by_role("cell", name="Grabar Buscar Editar Regresar").locator("button[name=\"btnGrabar\"]").click()
            time.sleep(2)
            print("Registro de clasificación completado exitosamente.")

            resultados_registro.append([curso_codigo[i][1],curso_codigo[i][0],"Clasificación Registrada Correctamente"])
            # ---------------------
        except Exception as e:
            print(f"Error durante el registro de clasificación: {e}")
            resultados_registro.append([curso_codigo[i][1],curso_codigo[i][0],"Clasificación No Registrada", str(e)])
            page.goto("https://ares.pucp.edu.pe/pucp/jsp/Intranet.jsp?url=%2Fpucp%2Fidiomas%2Fidwpanal%2FidwpanalPucpQuestionaccion%3DVerPanel")
            page.goto("https://ares.pucp.edu.pe/pucp/jsp/Intranet.jsp?url=%2Fpucp%2Fidiomas%2Fidwclasf%2FidwclasfPucpQuestionaccion%3DMostrarClasificacion")

        
    context.close()
    browser.close()
    print("Proceso de registro de clasificación finalizado.")
    print(resultados_registro)
    return resultados_registro

def fecha_y_hora_actual():
    # Obtener la fecha y hora actual
    fecha_hora_actual = datetime.now()
    # Imprimir la fecha y hora actual en formato legible
    fecha_hora_formato = fecha_hora_actual.strftime("%Y-%m-%d %H:%M:%S")
    return fecha_hora_formato

def diferencia_tiempo(inicio, fin):
    # Calcular la diferencia entre las dos fechas
    diferencia = fin - inicio
    return diferencia

def handler(event, context):
    print(fecha_y_hora_actual())
    # Parametrización de la URL y credenciales
    URL_PANDORA = "https://pandora.pucp.edu.pe/pucp/login?TARGET=https%3A%2F%2Feros.pucp.edu.pe%2Fpucp%2Fjsp%2FIntranet.jsp"
    
    body = json.loads(event["body"])

    headers = event["headers"]
    USERNAME = headers.get("usuario")
    PASSWORD = headers.get("clave")

    print("Datos recibidos")
    curso_codigo = body

    respuesta =[]

    with sync_playwright() as playwright:
        try:
            respuesta = funcion_registro_clasificacion(playwright,URL_PANDORA,USERNAME,PASSWORD,curso_codigo)
        except Exception as e:
            print("Error en registro de clasificación:", str(e))
            respuesta = {"Error en registro de clasificaciones. Detalle: ": str(e)}
            print(fecha_y_hora_actual())
            return {"statusCode": 500,"body": json.dumps(respuesta)}
        print("Respuesta:", respuesta) 
        print(fecha_y_hora_actual())
        return {"statusCode": 200,"body": json.dumps(respuesta)}
    
if __name__ == "__main__":
    evento_prueba = {
        "URL": "https://it5gtgotgirrgut5rnuvusnuz40nwuix.lambda-url.us-east-1.on.aws/",
        "body": json.dumps([
                ["Inglés Básico 2", "II169157", "Observación de prueba 1"]
            ]),
        "headers": {
            "Content-Type": "application/json",
            "usuario": "W0026391",
            "clave": "TyS.Idiomas_26_01"
        }
    }
    handler(evento_prueba, None)

