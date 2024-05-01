import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import save_file

class ToDoListApp:
    def __init__(self, root):
        # Initialization of the main application window
        self.root = root
        self.root.title("To-Do List")

        self.tasks = {}  # Dictionary for storing tasks

        # Adding a drop-down menu to select the day of the week
        self.label_day = tk.Label(root, text="Select day:")
        self.label_day.pack()

        self.days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        self.selected_day = tk.StringVar(value=self.days[0])

        self.combobox_day = ttk.Combobox(root, textvariable=self.selected_day, values=self.days, state='readonly')
        self.combobox_day.pack()

        # Other interface elements
        self.label_task = tk.Label(root, text="Enter your task:")
        self.label_task.pack()

        self.entry_task = tk.Entry(root)
        self.entry_task.pack()

        self.button_add = tk.Button(root, text="Add Task", command=self.add_task)
        self.button_add.pack()

        self.label_index = tk.Label(root, text="Enter task number to delete:")
        self.label_index.pack()

        self.entry_index = tk.Entry(root)
        self.entry_index.pack()

        self.button_delete = tk.Button(root, text="Delete Task", command=self.del_task)
        self.button_delete.pack()

        self.listbox_tasks = tk.Listbox(root, width=60)
        self.listbox_tasks.pack()

        self.button_export = tk.Button(root, text="Export to Excel", command=self.save_file)
        self.button_export.pack()

    def add_task(self):
        # Adding a task on the appropriate day
        day = self.selected_day.get()
        task = self.entry_task.get()

        if day not in self.tasks:
            self.tasks[day] = []
        self.tasks[day].append(task)
        self.update_tasks_list()
        messagebox.showinfo("Success", f"Task '{task}' added successfully.")

    def del_task(self):
        # Deleting a task by the specified index
        try:
            day = self.selected_day.get()
            task_index = int(self.entry_index.get()) - 1
            if task_index >= 0:
                tasks_for_day = self.tasks.get(day, [])
                if task_index < len(tasks_for_day):
                    deleted_task = tasks_for_day.pop(task_index)
                    self.update_tasks_list()
                    messagebox.showinfo("Success", f"Task '{deleted_task}' deleted successfully.")
                else:
                    messagebox.showerror("Error", "Invalid task index.")
            else:
                messagebox.showerror("Error", "Invalid task index.")
        except ValueError:
            messagebox.showerror("Error", "Invalid task index.")

    def update_tasks_list(self):
        # Updating the task list in the window
        self.listbox_tasks.delete(0, tk.END)
        for day in self.days:
            tasks_for_day = self.tasks.get(day, [])
            for index, task in enumerate(tasks_for_day, start=1):
                self.listbox_tasks.insert(tk.END, f"{index} {day}: {task}")

    def save_file(self):
        # Exporting tasks to Excel
        response = messagebox.askyesno("Export to Excel", "Do you want to export your to-do list to Excel?")
        if response:
            save_file.save_to_excel(self.tasks)


def main():
    # Creating and launching the main window
    root = tk.Tk()
    app = ToDoListApp(root)
    root.geometry("400x600")
    root.configure(bg="#bfbfbf")
    root.mainloop()


if __name__ == "__main__":
    main()
