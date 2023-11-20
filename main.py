from tkinter import messagebox
import subprocess
import sys
import tkinter as tk
import platform
from tkinter import ttk, filedialog
from ttkthemes import ThemedStyle
import os
import datetime
import struct
import csv
import serial
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.colors as mcolors

class DataCollectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("New Works Clearwell Level")
        self.root.geometry("1800x1600")

        self.current_time = datetime.datetime.now().strftime('%H:%M:%S')
        self.last_30_samples = []
        self.last_300_samples = []
        self.serial_status = "Not Connected"  # Initialize serial_status here
        self.error_messages = []  # New instance variable to store error messages



        # Bind the closing event to the custom function
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Use ttkthemes
        style = ThemedStyle(root)
        style.set_theme('plastik')  # You can change 'plastik' to other themes

        self.dark_mode = tk.BooleanVar()
        self.dark_mode.set(False)  # Default to light mode
        self.connected = False

        self.create_widgets()

    def create_widgets(self):

        # Labels
        self.current_time_label = tk.Label(self.root, text="Current Time: " + self.current_time, font=("Arial", 14))
        self.current_time_label.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)

        self.current_level_label = tk.Label(self.root, text="Current Level: Not Ready", font=("Arial", 14))
        self.current_level_label.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)

        self.status_label = tk.Label(self.root, text="Status: Not Ready", font=("Arial", 14))
        self.status_label.grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)

        # Serial Status
        self.serial_status_label = tk.Label(self.root, text="Serial Port: " + self.serial_status, font=("Arial", 14))
        self.serial_status_label.grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)

        # Combobox for selecting COM port


        com_ports = [f"COM{i}" for i in range(1, 21)]
        self.selected_com_port = tk.StringVar()
        self.com_port_combobox = ttk.Combobox(self.root, textvariable=self.selected_com_port, values=com_ports,
                                              state="readonly")
        self.selected_com_port.set("COM6")  # Set default value
        self.com_port_combobox.grid(row=11, column=0, padx=10, pady=10, sticky=tk.W)
        self.com_port_combobox.bind("<<ComboboxSelected>>", self.on_combobox_selected)

        # Combobox for selecting overflow level
        overflow_levels = [str(i) for i in range(1, 4)]
        self.selected_overflow_level = tk.StringVar()
        self.overflow_level_combobox = ttk.Combobox(self.root, textvariable=self.selected_overflow_level,
                                                    values=overflow_levels, state="readonly")
        self.selected_overflow_level.set("2")  # Set default value
        self.overflow_level_combobox.grid(row=12, column=0, padx=10, pady=10, sticky=tk.W)
        self.overflow_level_combobox.bind("<<ComboboxSelected>>", self.on_overflow_level_selected)

        # Combobox for selecting low level
        low_levels = [str(i) for i in range(8, 14)]
        self.selected_low_level = tk.StringVar()
        self.low_level_combobox = ttk.Combobox(self.root, textvariable=self.selected_low_level, values=low_levels,
                                               state="readonly")
        self.selected_low_level.set("8")  # Set default value
        self.low_level_combobox.grid(row=13, column=0, padx=10, pady=10, sticky=tk.W)
        self.low_level_combobox.bind("<<ComboboxSelected>>", self.on_low_level_selected)





        # Buttons
        self.connect_button = ttk.Button(self.root, text="Connect", command=self.connect_serial)
        self.connect_button.grid(row=9, column=0, padx=10, pady=10, sticky=tk.W)

        self.stop_button = ttk.Button(self.root, text="Disconnect", command=self.stop_serial)
        self.stop_button.grid(row=10, column=0, padx=10, pady=10, sticky=tk.W)

        self.open_file_button = ttk.Button(self.root, text="Open CSV File", command=self.open_file)
        self.open_file_button.grid(row=14, column=0, padx=10, pady=10, sticky=tk.W)

        self.clear_error_button = ttk.Button(self.root, text="Clear Errors", command=self.clear_error)
        self.clear_error_button.grid(row=18, column=5, padx=10, pady=10, sticky=tk.W)

        self.error_list_label = tk.Label(self.root, text="Error List", font=("Arial", 14))
        self.error_list_label.grid(row=16, column=2, padx=10, pady=10, sticky=tk.W)

        # Scrollbar for the error list
        self.error_scrollbar = tk.Scrollbar(self.root, orient=tk.VERTICAL)
        self.error_scrollbar.grid(row=18, column=3, sticky=tk.N + tk.S)

        # Error display
        self.error_display_var = tk.StringVar(value="")
        self.error_display = tk.Listbox(
            self.root, listvariable=self.error_display_var, selectmode=tk.SINGLE, width=180,
            yscrollcommand=self.error_scrollbar.set
        )
        self.error_display.grid(row=18, column=2, padx=10, pady=10, sticky=tk.W)

        # Configure the scrollbar to work with the error list
        self.error_scrollbar.config(command=self.error_display.yview)

        # Button to restart the application
        restart_button = ttk.Button(self.root, text="Restart", command=self.restart_application)
        restart_button.grid(row=19, column=0, padx=10, pady=10, sticky=tk.W)

        # Graph
        self.fig, self.ax = plt.subplots(figsize=(10, 5))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.root)
        self.canvas.get_tk_widget().grid(row=0, column=2, rowspan=17,columnspan=55, padx=10, pady=10, sticky=tk.E + tk.W + tk.N + tk.S)

        # Dark mode switch
        dark_mode_label = ttk.Label(self.root, text="Dark Mode:")
        dark_mode_label.grid(row=18, column=0, padx=10, pady=10, sticky=tk.W)

        dark_mode_switch = ttk.Checkbutton(self.root, variable=self.dark_mode, command=self.toggle_dark_mode)
        dark_mode_switch.grid(row=18, column=1, padx=10, pady=10, sticky=tk.W)


        # Serial connection
        self.ser = None

        # Initial plot
        self.plot_graph()

        # Select default values
        self.com_port_combobox.set(self.selected_com_port.get())
        self.overflow_level_combobox.set(self.selected_overflow_level.get())
        self.low_level_combobox.set(self.selected_low_level.get())

    def open_file(self):
        try:
            file_path = filedialog.askopenfilename(title="Select CSV file", filetypes=[("CSV files", "*.csv")])
            print("File Path:", file_path)  # Add this line for debugging
            if file_path:
                subprocess.run(["start", "", file_path], shell=True)
        except Exception as e:
            self.error_display.insert(tk.END, f"Error opening CSV file: {e}")

    def connect_serial(self):
        try:
            self.ser = serial.Serial(self.selected_com_port.get(), baudrate=9600, timeout=1)
            self.serial_status = "Connected"
            self.connected = True  # Set connected attribute to True
            self.update_status("Serial connected")
            self.root.after(10000, self.read_serial)
            # Schedule a function to check connection status every 5000 milliseconds (5 seconds)
            self.root.after(5000, self.check_connection_status)

        except serial.SerialException as se:
            self.serial_status = "Not Connected"
            self.connected = False  # Set connected attribute to False in case of an error
            self.error_display.insert(tk.END, f"Error opening serial port: {se}")

            # Update the serial status label
            self.serial_status_label.config(text="Serial Port: " + self.serial_status)

    def check_connection_status(self):
        try:
            if not self.ser.is_open:
                self.serial_status = "Not Connected"
                self.connected = False
                self.update_status("Serial disconnected")
                self.ser.close()  # Close the serial port if it's still open
                self.ser = None  # Set the serial port to None
            else:
                self.root.after(5000, self.check_connection_status)  # Reschedule the check

            # Update the serial status label
            self.serial_status_label.config(text="Serial Port: " + self.serial_status)

        except Exception as e:
            self.error_display.insert(tk.END, f"Error checking connection status: {e}")

    def stop_serial(self):
        try:
            if self.ser:  # Check if self.ser is not None before trying to close it
                self.ser.close()
                self.serial_status = "Not Connected"
                self.connected = False  # Set connected attribute to False when disconnected
                self.update_status("Serial disconnected")

        except Exception as e:
            self.error_display.insert(tk.END, f"Error closing serial port: {e}")

            # Update the serial status label
            self.serial_status_label.config(text="Serial Port: " + self.serial_status)



    def read_serial(self):
        try:
            data = self.ser.read(4)

            if data and len(data) == 4:
                steps = round(struct.unpack('f', data)[0], 1)
                self.update_time()
                self.update_level(steps)
                self.write_to_csv(steps)
                self.last_30_samples.append(steps)
                self.last_300_samples.append(steps)
                if len(self.last_300_samples)>300:
                    self.last_300_samples.pop(0)

                self.plot_graph()  # Call plot_graph immediately after updating the plot data
                if len(self.last_30_samples) > 30:
                    self.last_30_samples.pop(0)
                self.check_increase_or_decrease()
        except Exception as e:
            timestamp = datetime.datetime.now().strftime('%H:%M:%S')
            error_message = f"{timestamp} - Error: {e}"
            self.error_display.insert(tk.END, error_message)

        self.root.after(10000, self.read_serial)

    def write_to_csv(self, data):
        try:
            header = ['Date', 'Time', 'Level', 'Direction', 'Rate of Increase', 'Rate of Decrease', 'Time to Overflow',
                      'Time to Low Level', 'Minimum Level', 'Maximum Level']

            now = datetime.datetime.now()

            # Check if it's after 7 am, if not, use the previous day's CSV file
            if now.time() >= datetime.time(7, 0, 0):
                filename = 'cwwl_for_' + now.strftime('%Y-%m-%d') + '.csv'
            else:
                yesterday = now - datetime.timedelta(days=1)
                filename = 'cwwl_for_' + yesterday.strftime('%Y-%m-%d') + '.csv'

            file_exists = os.path.isfile(filename)

            date = now.strftime('%Y-%m-%d')
            time_str = now.strftime('%H:%M:%S')
            level = data

            status_text = self.status_label.cget("text")
            direction = status_text.split("\n")[2].split(":")[1].strip() if "Direction:" in status_text else ""
            rate_of_increase = status_text.split("\n")[3].split(":")[
                1].strip() if "Rate of Increase:" in status_text else ""
            rate_of_decrease = status_text.split("\n")[4].split(":")[
                1].strip() if "Rate of Decrease:" in status_text else ""
            time_to_overflow = status_text.split("\n")[5].split(":")[
                1].strip() if "Time to Overflow:" in status_text else ""
            time_to_low_level = status_text.split("\n")[6].split(":")[
                1].strip() if "Time to Low Level:" in status_text else ""
            minimum_level = self.selected_low_level.get() if self.selected_low_level.get() else "0"
            maximum_level = self.selected_overflow_level.get() if self.selected_overflow_level.get() else "0"

            with open(filename, 'a', newline='') as file:
                writer = csv.writer(file)
                if not file_exists:
                    writer.writerow(header)

                writer.writerow([date, time_str, level, direction, rate_of_increase, rate_of_decrease, time_to_overflow,
                                 time_to_low_level, minimum_level, maximum_level])

        except Exception as e:
            self.error_display.insert(tk.END, f"Error writing to CSV: {e}")

    def on_close(self):
        # Perform any cleanup or finalization here
        self.root.destroy()
        sys.exit()

    def plot_graph(self):
        try:
            if not self.connected:
                print("Not connected. Skipping plot.")
                return  # Do nothing if not connected

            print("Plotting graph...")
            self.ax.clear()

            # Customize the position and size of the graph
            right = 0.1  # Adjust the left margin
            bottom = 0.1  # Adjust the bottom margin
            width = 0.8  # Adjust the width
            height = 0.6  # Adjust the height
            self.ax.set_position([right, bottom, width, height])

            x = range(1, len(self.last_300_samples) + 1)
            y = self.last_300_samples[-300:]

            # Get color from calculate_fill_color function
            color = self.calculate_fill_color(y[-1])

            # Customize the color and style of the graph
            self.ax.plot(x, y, marker='o', linestyle='-', color=color)

            self.ax.fill_between(x, y, y2=20, color=color, alpha=0.8)

            # Customize x and y axis labels
            self.ax.set_xlabel('Sample', fontsize=12, color='blue')
            self.ax.set_ylabel('Level', fontsize=12, color='blue')

            # Add grid lines
            self.ax.grid(True, linestyle='--', alpha=0.7)

            # Customize tick labels
            self.ax.tick_params(axis='both', labelsize=10, colors='black')

            # Customize the y-axis range
            self.ax.set_ylim(0, 20)  # Adjust the range as needed

            # Set yticks to have a step of 1
            self.ax.set_yticks(range(0, 21))

            self.ax.invert_yaxis()
            print(f"Level: {y[-1]}, Red Intensity: {min(int((y[-1] - 1) * 10), 255)}")

            self.canvas.draw()
            print("Graph plotted successfully.")


        except Exception as e:
            self.error_display.insert(tk.END, f"Error plotting graph: {e}")

    def check_increase_or_decrease(self):
        try:
            if not self.connected:
                print("Not connected. Skipping check increase/decrease.")
                return  # Do nothing if not connected

            print("Checking increase or decrease...")
            current_time = datetime.datetime.now().strftime('%H:%M:%S')

            if len(self.last_30_samples) < 30:
                self.status_label.config(
                    text=f"Current Time: {current_time}\nCurrent Level: {self.last_30_samples[-1]}\nDirection: Not Ready\nRate of Increase: Not Ready\nRate of Decrease: Not Ready\nTime to Overflow: Not Ready\nTime to Low Level: Not Ready")
            else:
                current_level = self.last_30_samples[-1]
                total_difference = self.last_30_samples[-1] - self.last_30_samples[0]

                # Get minimum and maximum levels from the drop-down menus
                minimum_level = int(self.selected_low_level.get())
                maximum_level = int(self.selected_overflow_level.get())

                rate_of_increase, rate_of_decrease = 0, 0
                estimated_overflow_time, estimated_low_level_time = "Not Ready", "Not Ready"

                if total_difference > 0:
                    direction = "decreasing"
                    rate_of_decrease = abs(total_difference * 3600) / (300)  # Rate of decrease in ft/hour

                    # Calculate time to reach minimum
                    time_to_reach_min = (minimum_level - current_level) / rate_of_decrease
                    # time_to_reach_min_formatted = str(datetime.timedelta(seconds=int(time_to_reach_min * 3600)))
                    estimated_low_level_time = ( datetime.datetime.now() + datetime.timedelta(seconds=int(time_to_reach_min * 3600))).strftime(
                        '%H:%M:%S')

                elif total_difference < 0:
                    direction = "increasing"
                    rate_of_increase = abs(total_difference * 3600) / (300)  # Rate of increase in ft/hour
                    time_to_reach_max = (current_level - maximum_level) / rate_of_increase
                    # time_to_reach_max_formatted = str(datetime.timedelta(seconds=int(time_to_reach_max * 3600)))
                    estimated_overflow_time = (
                            datetime.datetime.now() + datetime.timedelta(seconds=int(time_to_reach_max * 3600))).strftime(
                        '%H:%M:%S')

                else:
                    direction, rate_of_increase, rate_of_decrease, estimated_overflow_time, estimated_low_level_time = "stagnant", 0, 0, "Not Ready", "Not Ready"

                self.status_label.config(
                    text=f"Current Time: {current_time}\nCurrent Level: {current_level}\nDirection: {direction}\nRate of Increase: {rate_of_increase:.2f} ft/hour\nRate of Decrease: {rate_of_decrease:.2f} ft/hour\nTime to Overflow: {estimated_overflow_time}\nTime to Low Level: {estimated_low_level_time}")
                print("Checked increase or decrease successfully.")

        except Exception as e:
            self.error_display.insert(tk.END, f"Error checking increase or decrease: {e}")

    def update_time(self):
        self.current_time = datetime.datetime.now().strftime('%H:%M:%S')
        self.current_time_label.config(text="Current Time: " + self.current_time)

    def update_level(self, level):
        self.current_level_label.config(text="Current Level: " + str(level))

    def update_status(self, message):
        self.status_label.config(text="Status: " + message)

    def restart_application(self):
        confirm_restart = messagebox.askyesno("Confirmation", "Are you sure you want to restart the application?")
        if confirm_restart:
            self.root.destroy()
            # Restart the application by creating a new Tkinter root window
            root = tk.Tk()
            app = DataCollectorApp(root)
            root.option_add("*tearOff", False)
            root.mainloop()

    def log_error(self, message):
        timestamp = datetime.datetime.now().strftime('%H:%M:%S')
        error_message = f"[{timestamp}] {message}"
        self.error_display_var.set(self.error_display_var.get() + "\n" + error_message)
        self.error_display.see(tk.END)  # Scroll to the end to show the latest error


    def clear_error(self):
        self.error_messages.clear()
        self.error_display_var.set("")

    def toggle_dark_mode(self):
        if self.mode == "light":
            self.configure(bg="darkblue")
            self.toggle_button.configure(style="TButton")
            self.style.configure("TButton", foreground="white", background="darkblue")
            self.mode = "dark"
        else:
            self.configure(bg="white")
            self.toggle_button.configure(style="TButton")
            self.style.configure("TButton", foreground="black", background="white")
            self.mode = "light"

    def on_combobox_selected(self, event):
        self.selected_com_port = self.com_port_combobox.get()
        self.update_status(f"Selected COM Port: {self.selected_com_port}")
        # Update minimum and maximum levels based on selected overflow and low levels
        self.minimum_level = int(self.selected_low_level.get())
        self.maximum_level = int(self.selected_overflow_level.get())
        self.check_increase_or_decrease()  # Recalculate values based on the new levels

    def on_overflow_level_selected(self, event):
        self.update_status(f"Selected Overflow Level: {self.selected_overflow_level.get()}")
        self.maximum_level = int(self.selected_overflow_level.get())

    def on_low_level_selected(self, event):
        self.update_status(f"Selected Low Level: {self.selected_low_level.get()}")
        self.minimum_level = int(self.selected_low_level.get())




    def calculate_fill_color(self, level):
        minimum = float(self.selected_low_level.get())
        maximum = float(self.selected_overflow_level.get())
        near_minimum = minimum - 1
        near_maximum = maximum + 1

        # Create a custom colormap for the color intensity changes
        colors = ['orange', 'red']
        cmap = mcolors.LinearSegmentedColormap.from_list("", colors)

        if near_minimum < level < minimum or near_maximum > level > maximum:
            # Calculate the color intensity based on the level
            if near_minimum < level < minimum:
                intensity = level - near_minimum
            else:  # near_maximum > level > maximum
                intensity = near_maximum - level
            return cmap(intensity)
        elif level == near_minimum or level == near_maximum:
            return 'orange'
        elif near_minimum> level >near_maximum:
            return 'lightblue'
        else:
            return 'red'

    def open_file(self):
        try:
            file_path = filedialog.askopenfilename(title="Select CSV file", filetypes=[("CSV files", "*.csv")])
            print("File Path:", file_path)  # Add this line for debugging
            if file_path:
                # Use platform-specific command to open file with default application
                if platform.system() == 'Windows':
                    subprocess.run(["start", "excel.exe", file_path], shell=True)
                elif platform.system() == 'Darwin':
                    subprocess.run(["open", "-a", "Microsoft Excel", file_path])
                elif platform.system() == 'Linux':
                    subprocess.run(["xdg-open", file_path])
        except Exception as e:
            self.error_display.insert(tk.END, f"Error opening CSV file: {e}")





if __name__ == "__main__":
    root = tk.Tk()
    app = DataCollectorApp(root)
    root.option_add("*tearOff", False)  # This is always a good idea

    root.mainloop()
