import win32com.client
import tkinter as tk
from tkinter import ttk

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
            tasks.append((task.Name, task.NextRunTime, status, parameters))

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
            params += f"{action.Path} {action.Arguments}\n"
    return params.strip()

def copy_selected(event=None):
    selected_items = treeview_tasks.selection()
    if selected_items:
        copied_data = ""
        for item in selected_items:
            copied_data += "\t".join(treeview_tasks.item(item, 'values')) + "\n"
        window.clipboard_clear()
        window.clipboard_append(copied_data)

def execute_selected(event):
    selected_item = treeview_tasks.focus()
    if selected_item:
        task_name = treeview_tasks.item(selected_item, 'values')[0]
        print(f"Ejecutando tarea: {task_name}")
        # Aquí puedes agregar la lógica para ejecutar la tarea seleccionada

def search_tasks(event=None):
    # Obtener el texto de búsqueda de la entrada de texto
    search_text = entry_search.get()

    # Limpiar el contenido del Treeview
    treeview_tasks.delete(*treeview_tasks.get_children())

    # Obtener las tareas programadas
    scheduled_tasks = get_scheduled_tasks(entry_server.get(), search_text)
    if len(scheduled_tasks) > 0:
        for task_name, next_run_time, status, parameters in scheduled_tasks:
            treeview_tasks.insert('', tk.END, values=(task_name, next_run_time, status, parameters))
    else:
        treeview_tasks.insert('', tk.END, values=("No se encontraron tareas que coincidan con el texto de búsqueda.", "", "", ""))

# Configuración del servidor
default_server_name = 'SERVERSAP'

# Crear la ventana principal
window = tk.Tk()
window.title("Búsqueda de Tareas Programadas")

# Crear etiqueta y entrada de texto para el servidor
label_server = tk.Label(window, text="Servidor:")
label_server.pack()

entry_server = tk.Entry(window)
entry_server.pack()
entry_server.insert(0, default_server_name)

# Crear etiqueta y entrada de texto para el texto de búsqueda
label_search = tk.Label(window, text="Texto de Búsqueda:")
label_search.pack()

entry_search = tk.Entry(window)
entry_search.pack()

# Vincular la tecla "Enter" a la función search_tasks
entry_search.bind('<Return>', search_tasks)

# Crear el Treeview para mostrar las tareas en un grid
treeview_tasks = ttk.Treeview(window, columns=('Task', 'Next Run Time', 'Status', 'Parameters'), show='headings')
treeview_tasks.heading('Task', text='Tarea')
treeview_tasks.heading('Next Run Time', text='Próxima ejecución')
treeview_tasks.heading('Status', text='Estado')
treeview_tasks.heading('Parameters', text='Parámetros')
treeview_tasks.pack()

# Ajustar el ancho de las columnas
treeview_tasks.column('Task', width=200)
treeview_tasks.column('Next Run Time', width=200)
treeview_tasks.column('Status', width=100)
treeview_tasks.column('Parameters', width=800)

# Configurar el alto del Treeview
treeview_tasks.configure(height=15)

# Configurar menú contextual para copiar
context_menu = tk.Menu(window, tearoff=0)
context_menu.add_command(label="Copiar", command=copy_selected)

# Vincular el menú contextual al Treeview
treeview_tasks.bind("<Button-3>", lambda event: context_menu.post(event.x_root, event.y_root))

# Configurar evento de doble clic para ejecutar la tarea
treeview_tasks.bind("<Double-1>", execute_selected)

# Obtener la anchura y altura de la pantalla
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()

# Calcular las nuevas dimensiones de la ventana
window_width = int(screen_width * 0.7)  # Utilizamos el 70% del ancho de la pantalla
window_height = int(screen_height * 0.7)  # Utilizamos el 70% de la altura de la pantalla

# Configurar el tamaño de la ventana
window.geometry(f"{window_width}x{window_height}")

# Iniciar el bucle principal de la interfaz
window.mainloop()

