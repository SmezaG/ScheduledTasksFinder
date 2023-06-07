import win32com.client
import tkinter as tk
from tkinter import ttk
import subprocess
from cryptography.fernet import Fernet
import configparser
from PIL import ImageTk, Image 
import datetime


def get_scheduled_tasks(server_name, search_text):
    tasks = []
    scheduler = win32com.client.Dispatch(f'Schedule.Service.1')
    scheduler.Connect(server_name)
    folder = scheduler.GetFolder('\\')
    task_collection = folder.GetTasks(0)
    num_tasks = task_collection.Count

    for i in range(1, num_tasks + 1):
        task = task_collection.Item(i)
        if search_text.lower() in task.Name.lower():
            status = get_task_status(task)
            parameters = get_task_parameters(task)
            arguments = get_task_arguments(task)
            LastExcution = task.LastRunTime.strftime("%d/%m/%Y %H:%M:%S")
            tasks.append((task.Name, LastExcution, status, parameters,arguments))

    return tasks

def get_task_status(task):
    if task.Enabled:
        if task.State == 4:
            return 'En ejecución'
        else:
            return 'Lista'
    else:
        return 'Deshabilitada'

def get_task_parameters(task):
    definition = task.Definition
    params = ""
    for action in definition.Actions:
        if action.Type == 0:  # Script Action
            params += f"{action.Path}\n"
    return params.strip()

def get_task_arguments(task):
    definition = task.Definition
    params = ""
    for action in definition.Actions:
        if action.Type == 0:  # Script Action
            params += f"{action.Arguments}\n"
    return params.strip()

def copy_selected(event=None):
    selected_items = treeview_tasks.selection()
    if selected_items:
        copied_data = ""
        selected_cell = treeview_tasks.focus()
        values = treeview_tasks.item(selected_cell, 'values')
        if values:
            copied_data = values[4]
        window.clipboard_clear()
        window.clipboard_append(copied_data)

def decrypt_ini_file(file_path, key):
    with open(file_path, "rb") as file:
        file_data = file.read()
    fernet = Fernet(key)
    decrypted_data = fernet.decrypt(file_data)
    return decrypted_data.decode()

def execute_selected(event=None):
    selected_item = treeview_tasks.focus()
    if selected_item:
        task_name = treeview_tasks.item(selected_item, 'values')[0]
        print(f"Ejecutando tarea: {task_name}")

        # Comando para ejecutar la tarea en el servidor remoto con el usuario "username"
        command = [
            "schtasks",
            "/run",
            "/s",
            server_name,
            "/tn",
            task_name,
            "/u",
            username,
            "/p",
            password
        ]
        subprocess.run(command, shell=True)

         # Actualizar la columna de estado en el treeview
        treeview_tasks.set(selected_item, 'Status', 'En ejecución')
        

def search_tasks(event=None):
    # Obtener el texto de búsqueda de la entrada de texto
    search_text = entry_search.get()

    # Limpiar el contenido del Treeview
    treeview_tasks.delete(*treeview_tasks.get_children())

    # Obtener las tareas programadas
    scheduled_tasks = get_scheduled_tasks(entry_server.get(), search_text)
    if len(scheduled_tasks) > 0:
        for task_name, next_run_time, status, parameters, Arguments in scheduled_tasks:
            treeview_tasks.insert('', tk.END, values=(task_name, next_run_time, status, parameters, Arguments), tags=("Data",))
    else:
        treeview_tasks.insert('', tk.END, values=("No se encontraron tareas que coincidan con el texto de búsqueda.", "", "", ""))

def TreeviewCreator(window):
    treeview_tasks = ttk.Treeview(window, columns=('Task', 'Trigger', 'Status', 'Parameters','Arguments'), show='headings')
    treeview_tasks.heading('Task', text='Tarea', command=lambda: sort_column(treeview_tasks, 'Task'))
    treeview_tasks.heading('Trigger', text='Última ejecución', command=lambda: sort_column(treeview_tasks, 'Trigger'))
    treeview_tasks.heading('Status', text='Estado', command=lambda: sort_column(treeview_tasks, 'Status'))
    treeview_tasks.heading('Parameters', text='Ruta', command=lambda: sort_column(treeview_tasks, 'Parameters'))
    treeview_tasks.heading('Arguments', text='Parámetros', command=lambda: sort_column(treeview_tasks, 'Arguments'))
    #treeview_tasks.pack()

    # Ajustar el ancho de las columnas
    treeview_tasks.column('Task', width=200)
    treeview_tasks.column('Trigger', width=200)
    treeview_tasks.column('Status', width=100)
    treeview_tasks.column('Parameters', width=600)
    treeview_tasks.column('Arguments', width=200)
    return treeview_tasks

def on_key_release(event=None):

    # Si pulsamos Enter buscará, si no borrará el contenido
    if event.keysym == "Return" or event.keysym == "F5":
        search_tasks()
    elif not entry_search.get():
         treeview_tasks.delete(*treeview_tasks.get_children())

def Update_task_status(event=None):
    selected_item = treeview_tasks.focus()
    if selected_item:

        task_name = treeview_tasks.item(selected_item, 'values')[0]
        current_status = treeview_tasks.item(selected_item, 'values')[2]
        if current_status == 'Deshabilitada':
            new_status = 'Lista'
        elif current_status == 'Lista':
            new_status = 'Deshabilitada'
    if new_status:
            treeview_tasks.set(selected_item, 'Status', new_status)
            scheduler = win32com.client.Dispatch(f'Schedule.Service.1')
            command = [
            "schtasks",
            "/change",
            "/s",
            server_name,
            "/tn",
            task_name,
            "/u",
            username,
            "/p",
            password
        ]
            if new_status == 'Lista':
                command.extend(["/enable"])
            elif new_status == 'Deshabilitada':
                command.extend(["/disable"])
            subprocess.run(command, shell=True)

def DecriptorMap():
    # Cargar y desencriptar el archivo .ini

        global server_name, username, password

        encrypted_file_path = "params.ini"
        keyfile = open("key.txt","r")
        encryption_key = keyfile.read()
        keyfile.close()
        decrypted_ini_data = decrypt_ini_file(encrypted_file_path, encryption_key)

        config_parser = configparser.ConfigParser()
        config_parser.read_string(decrypted_ini_data)

        server_name = config_parser.get("Credentials", "servidor")
        username = config_parser.get("Credentials", "usuario")
        password = config_parser.get("Credentials", "password")

def sort_column(treeview, col, reverse=False):
    # Obtiene todos los elementos del Treeview
    data = [(treeview.set(child, col), child) for child in treeview.get_children('')]

    # Ordena los elementos en función del valor de la columna
    data.sort(reverse=reverse)

    for index, (value, child) in enumerate(data):
        # Reorganiza los elementos en el Treeview
        treeview.move(child, '', index)

    # Cambia la dirección de ordenamiento para el próximo clic en el encabezado
    treeview.heading(col, command=lambda: sort_column(treeview, col, not reverse))


server_name = ""
username = ""
password = ""

DecriptorMap()

# Configuración del servidor
default_server_name = server_name

# Crear la ventana principal
window = tk.Tk()
window.title("Búsqueda de Tareas Programadas")

# Cargar la imagen del logo
logo_image = Image.open("logo.png")  # Reemplaza "logo.png" con la ruta y el nombre de tu archivo de imagen
logo_image = logo_image.resize((16, 16))  # Ajusta el tamaño del logo según tus necesidades

# Crear una instancia de la clase PhotoImage
window.iconphoto(True, ImageTk.PhotoImage(logo_image))

# Crear el estilo para los widgets
style = ttk.Style()
style.configure("TLabel", font=("Arial", 10))
style.configure("TEntry", font=("Arial", 10), borderwidth=0, relief="solid", padding=5)
style.configure("Treeview.Heading", font=("Arial", 10,))

# Obtener las dimensiones de la pantalla
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()

# Calcular las nuevas dimensiones de la ventana
window_width = int(screen_width * 0.7)
window_height = int(screen_height * 0.7)

# Configurar el tamaño y posición de la ventana
window.geometry(f"{window_width}x{window_height}+{int((screen_width-window_width)/2)}+{int((screen_height-window_height)/2)}")

# Crear el frame principal
frame_main = ttk.Frame(window, padding=20)
frame_main.pack(fill='both', expand=True)

# Crear etiqueta y entrada de texto para el servidor
label_server = ttk.Label(frame_main, text="Servidor:")
label_server.grid(row=0, column=0, sticky="e")

entry_server = ttk.Entry(frame_main, width=30)
entry_server.grid(row=0, column=1, padx=10, pady=5, sticky="w")
entry_server.insert(0, default_server_name)
entry_server.config(state = "readonly")

# Crear etiqueta y entrada de texto para el texto de búsqueda
label_search = ttk.Label(frame_main, text="Texto de Búsqueda:")
label_search.grid(row=1, column=0, sticky="e")

entry_search = ttk.Entry(frame_main, width=30)
entry_search.grid(row=1, column=1, padx=10, pady=5, sticky="w")

# Espacio en blanco
empty_label = ttk.Label(frame_main, text="")
empty_label.grid(row=2, column=0, columnspan=2)

# Vincular la tecla "Enter" a la función search_tasks
entry_search.bind('<KeyRelease>', on_key_release)

# Crear Treeview con estilo
treeview_tasks = TreeviewCreator(frame_main)
treeview_tasks.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
frame_main.grid_rowconfigure(3, weight=1)  # Ajustar el tamaño del treeview verticalmente
frame_main.grid_columnconfigure(0, weight=1)  # Ajustar el tamaño del treeview horizontalmente
frame_main.grid_columnconfigure(1, weight=1)  # Ajustar el tamaño del treeview horizontalmente


# Configurar menú contextual para copiar
context_menu = tk.Menu(window, tearoff=0)
context_menu.add_command(label="Copiar", command=copy_selected)
context_menu.add_command(label="Ejecutar", command=execute_selected)
context_menu.add_command(label="Habilitar/Deshabilitar", command=Update_task_status)

# Vincular el menú contextual al Treeview
treeview_tasks.bind("<Button-3>", lambda event: context_menu.post(event.x_root, event.y_root))

style.map("Treeview", bordercolor=[("active", "#0078D7")])

# Iniciar el bucle principal de la interfaz
window.mainloop()
