import iesve
import tkinter as tk
import xlsxwriter
import os
import numpy as np
def generate_window(project):

    class Window(tk.Frame):
        def __init__(self, master=None):
            tk.Frame.__init__(self, master)
            self.project = project
            self.project_folder = project.path
            self.save_file_name = 'Conduction_Gain_Data'
            self.master = master
            self.init_window()

        # Creation of init_window
        def init_window(self):

            # changing the title of our master widget
            
            self.master.title("Automated Energy Analysis")
            self.master.columnconfigure(0, weight=1)
            self.master.rowconfigure(0, weight=1)
            self.master.grid()

            #print(self.project_folder)
            tk.Label(self, text=' ').grid(row=7, sticky=tk.W)

            tk.Label(self, text='Conduction Gain Data will be added to an Excel sheet that will be saved into the main project folder').grid(row=8, sticky=tk.W)
            tk.Label(self, text='Name the Excel file below:').grid(row=9, sticky=tk.W)
            self.save_file_entry_box = tk.Entry(self)
            self.save_file_entry_box.insert(0, self.save_file_name)
            self.save_file_entry_box.grid(row=10, sticky='ew')
            tk.Label(self, text=' ').grid(row=11, sticky=tk.W)
            # creating a button instance
            tk.Button(self, text="Run Calculation", command=self.run_calc).grid(row=16, sticky=tk.W)
            
            self.columnconfigure(0, weight=1)
            self.grid(row=0, column=0, sticky='nsew')

        def run_calc(self):
            self.save_file_name = self.save_file_entry_box.get()
            print('Excel File name = \t\t' + self.save_file_name)
            
            # create excel workbook
            workbook = xlsxwriter.Workbook(self.project_folder + '\\' + self.save_file_name + '.xlsx')
            # create excel work sheet
            sheet1 = workbook.add_worksheet('sheet1')
            def get_conduction_gain():
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
                result = sim.run_simulation(batch_operation)

                filename = "../vista/results.aps"

                aps_file = iesve.ResultsReader.open(filename)

                project  = iesve.VEProject.get_current_project()

                room_ids = aps_file.get_room_ids()
                room_names = aps_file.get_room_list()
				
                data=[]
                for id in range(len(room_ids)): 
                    room_name = room_names[id][0]
                    room_data=["Conduction gain - external walls & Conduction gain - roof & Conduction gain - ground floor",room_name,"results.aps"]
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
                return data
                
          
            # run main calculation functions
            print('Running Calculations')
            shape_data = get_conduction_gain()
            heading = ['Var. Name',
                      'Building Name',
                       'Filename',
                       'Max. Val.',
                       'Min. Val.'
                      ]
								
            # write data to excel worksheets
            print('Writing results to Excel Sheet')

            # write results data
            row = 1
            filename = "../vista/results.aps"
            aps_file = iesve.ResultsReader.open(filename)
            room_ids = aps_file.get_room_ids()
            sheet1.write_row(row-1, 0, heading)
            for shape in range(len(room_ids)):
                sheet1.write_row(row,   0, shape_data[shape])
                row+=1
            
           
            try:
                workbook.close()
            except PermissionError as e:
                print("Couldn't close workbook: ", e)
            os.startfile(self.project_folder + '\\' + self.save_file_name + '.xlsx')
            root.destroy()

    root = tk.Tk()
    app = Window(root)
    root.mainloop()

if __name__ == '__main__':
    project = iesve.VEProject.get_current_project()

    generate_window(project)
