import iesve
import tkinter as tk
from pathlib import Path
import xlsxwriter
import os
import numpy as np
from math import cos, sin, tan
from tkinter.filedialog import askdirectory

def import_building(filepath):
    print(filepath)
    iesve.ImportGBXML.import_file(filepath,True,iesve.VolumeCapMode.none,0.0)
 
def get_conduction_gain(building_name):
    sim = iesve.ApacheSim()
    batch_operation = False
    sim.set_options({'results_filename':'results.aps',
    'output_conduction_gains':True,
     
    'output_HVAC_components':True,

    'output_HVAC_systems':True,

    'output_latent_internal_gains':True,

    'output_latent_ventilation_gains':True,

    'output_sensible_internal_gains':True,

    'output_standard_outputs':True

    })
    result   = sim.run_simulation(batch_operation)

    filename = "../vista/results.aps"

    aps_file = iesve.ResultsReader.open(filename)

    project  = iesve.VEProject.get_current_project()

    room_ids = aps_file.get_room_ids()
    room_names = aps_file.get_room_list()
    
    data=[]
    for id in range(len(room_ids)): 
        room_name = room_names[id][0]
        room_data=["Conduction gain",building_name]
        walls_gain = aps_file.get_all_room_results(room_ids[id],'Conduction gain - external wall','z',1,365)['Conduction gain - external walls']

        roof_gain  = aps_file.get_all_room_results(room_ids[id],'Conduction gain - roof','z',1,365)['Conduction gain - roofs']                 
        
        floor_gain = aps_file.get_all_room_results(room_ids[id],'Conduction gain - ground floor','z',1,365)['Conduction gain - ground/exposed floors']
        
        combined_var=[]

        for hour in range(0,8759):
            combined_var.append(walls_gain[hour]+roof_gain[hour]+floor_gain[hour])

        max = np.max(combined_var)
        min = np.min(combined_var)
        aps_units = aps_file.get_units()
        gain_units = aps_units['Gain']
        pu_ip = gain_units['units_metric']
        max_IP = ((max / pu_ip['divisor']) + pu_ip['offset'])
        gain_units = aps_units['Gain']
        pu_ip = gain_units['units_metric']
        min_IP = ((min / pu_ip['divisor']) + pu_ip['offset'])
        room_data.append('{} {}'.format(max_IP, pu_ip['display_name']))
        room_data.append('{} {}'.format(min_IP, pu_ip['display_name']))
        data.append(room_data)
    recent_data = []
    recent_data.append(data[-1])
    return recent_data
                
def generate_window(project):

    class Window(tk.Frame):
        def __init__(self, master=None):
            tk.Frame.__init__(self, master)
            self.project = project
            self.project_folder = project.path
            self.save_file_name = 'Conduction_Gain_Results'
            self.master = master
            self.init_window()

        # Creation of init_window
        def init_window(self):
            """
            Initializes the window with labels, entry boxes, and buttons.
            """
            # Change the title of the main window
            self.master.title("Conduction Gain Simulation")
            self.master.columnconfigure(0, weight=1)
            self.master.rowconfigure(0, weight=1)
            # Add this line inside the _init_ method
            self.imported_body_deleted = tk.BooleanVar()
            self.imported_body_deleted.set(False)
            self.imported_body_deleted.trace("w", lambda *args: self.check_body_deleted())
            empty_label_rows = [7, 11]
            for row in empty_label_rows:
                tk.Label(self, text=' ').grid(row=row, sticky=tk.W)

            tk.Label(self, text='Conduction Gain Data will be added to an Excel sheet that will be saved into the main project folder').grid(row=8, sticky=tk.W)
            tk.Label(self, text='Name the Excel file below:').grid(row=9, sticky=tk.W)

            self.save_file_entry_box = tk.Entry(self)
            self.save_file_entry_box.insert(0, self.save_file_name)
            self.save_file_entry_box.grid(row=10, sticky='ew')

            # Create button
            run_button = tk.Button(self, text="Run Calculation", command=self.run_process)
            run_button.grid(row=16, sticky=tk.W)

            self.columnconfigure(0, weight=1)
            self.grid(row=0, column=0, sticky=tk.NSEW)
            
        def body_deleted(self):
            room_count = -1
            project = iesve.VEProject.get_current_project()
            models = project.models
            while room_count != 0:
                room_count = 0
                for model in models:
                    try:
                        bodies = model.get_bodies_and_ids(False)
                        for id, body in bodies.items():
                            # We only want to process thermal rooms here, so filter by type
                            if body.type == iesve.VEBody_type.room:
                               body.select()
                               room_count += 1
                    except RuntimeError:
                        room_count = 0
                        break

                if room_count == 0:
                    break

            return room_count == 0      
        
        def check_body_deleted(self):
            if not self.body_deleted():
               self.master.after(1000, self.check_body_deleted)  # Schedule to check again after 1000 ms (1 second)
                
        def close_main_window(self):
            self.master.withdraw()
            
        def run_process(self):
            self.save_file_name = self.save_file_entry_box.get()
            print('Excel File name = \t\t' + self.save_file_name)
            # create excel workbook
            workbook = xlsxwriter.Workbook(self.project_folder + '\\' + self.save_file_name + '.xlsx')
            # create excel work sheet
            sheet1 = workbook.add_worksheet('sheet1')
            
            def import_delete(self,foldername):
                data = []
                filenames = os.listdir(foldername)
                for filename in filenames:
                    f = os.path.join(foldername, filename)
                    if os.path.isfile(f):
                        import_path = f
                        import_building(import_path)
                        data += get_conduction_gain(os.path.splitext(os.path.basename(filename))[0])
                        self.check_body_deleted()       
                return data
                
            # run main calculation functions
            print('Running Calculations')
            foldername = askdirectory()
            self.close_main_window()
            buildings_data = import_delete(self,foldername)
            heading    = [
                       'Var. Name',
                       'Building Name',
                       'Max. Val.',
                       'Min. Val.'
                         ]
								
            # write data to excel worksheets
            print('Writing results to Excel Sheet')

            # write results data
            y = 1

            sheet1.write_row(y-1, 0, heading)
            for buildings_data in buildings_data:
                sheet1.write_row(y,   0, buildings_data)
                y+=1
           
            try:
                workbook.close()
            except PermissionError as e:
                print("Couldn't close workbook: ", e)
            os.startfile(self.project_folder + '\\' + self.save_file_name + '.xlsx')

    root = tk.Tk()
    app = Window(root)
    root.mainloop()

if __name__ == '__main__':
    project = iesve.VEProject.get_current_project()

    generate_window(project)