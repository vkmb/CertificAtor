"""
Application for writing e-cetrificates
"""
import os
import csv
import toga
# import cv2
from toga.style import Pack
from threading import Thread
from PIL import Image, ImageDraw, ImageFont
from toga.style.pack import COLUMN, ROW, NONE

class CertificAtor(toga.App):

    def startup(self):
        """
        Construct and show the Toga application.

        Usually, you would add your application to a main content box.
        We then create a main window (with a name matching the app), and
        show the main window.
        """
        self.file_name = ""
        self.font_file_name = ""
        self.tmp_filen = ""
        self.font_pil_ttf = None
        self.width, self.height = 1000, 1000
        
        self.csv_file_name = toga.sources.ValueSource()
        self.loaded_csv_data = toga.sources.ListSource([], [])
        
        self.coord = {}
        self.bnd_box = 0
        self.text_list = []
        self.box_list = []
        self.x_in = []
        self.y_in = []

        self.left_panel = toga.Box(style=Pack(direction=COLUMN))        
        self.certificate_box = toga.Box(style=Pack(direction=ROW, padding=5))
        file_name_label = toga.Label("Certificate Image Name", style=Pack(padding=(0, 5)))
        self.file_name_input = toga.TextInput(style=Pack(flex=1))
        self.font_label_loc = toga.Label("Font File : None", style=Pack(padding=5))
        self.fetch_button = toga.Button("Open", id="open", style=Pack(padding=(0, 5)), on_press=self.fetch)
        self.loading = toga.ProgressBar(style=Pack(padding=5))
        self.certificate_box.add(file_name_label, self.file_name_input, self.fetch_button)
        
        self.name_container = toga.Box(style=Pack(direction=ROW, padding=5))
        self.global_ctrl_box  = toga.Box(style=Pack(direction=COLUMN, padding=5))
        
        self.add_name_button = toga.Button("Add Label", id="add", style=Pack(padding=5, flex=1), on_press=self.fetch)
        self.add_font_file = toga.Button("Change Font", id="font", style=Pack(padding=5, flex=1), on_press=self.fetch)
        self.f = toga.Label('Font Size :', id=f'font_size_label', style=Pack(padding=5))
        self.f_in = toga.Slider(id=f'fin', style=Pack(padding=5), default=10, range=(10, 100), on_slide=self.teleport)
        
        self.global_ctrl_box.add(self.add_name_button, self.add_font_file, self.f, self.f_in)

        self.name_box = toga.Box(style=Pack(direction=COLUMN, padding=5))
        self.name_scroller = toga.ScrollContainer(style=Pack(flex=1))
        self.name_scroller.content = self.name_box
        self.name_container.add(self.global_ctrl_box)
        self.name_container.add(self.name_scroller)
        self.image_view = toga.ImageView(style=Pack(flex=1, padding=5))
        self.left_panel.add(self.image_view, self.loading, self.name_container, self.certificate_box, self.font_label_loc)
        self.main_panel = toga.SplitContainer(style=Pack(flex=1))
        self.right_panel = toga.Box(style=Pack(flex=1, direction=COLUMN))
        self.open_csv = toga.Button("Open CSV File", id="open_csv", on_press=self.fetch)
        self.csv_widget = toga.Table(id="sheet", headings=[], style=Pack(flex=1, padding=5), on_select=self.teleport)
        self.right_panel.add(self.csv_widget)
        self.right_panel.add(self.open_csv)
        self.main_panel.content = [self.left_panel, self.right_panel]
        self.main_window = toga.MainWindow(title=self.formal_name)
        self.main_window.content = self.main_panel
        self.main_window.show()
    
    def fetch(self, widget):
        if widget.id == "open" or (widget.id == "add" and self.file_name == ""):
            if len(self.file_name_input.value) == 0:
                self.file_name_input.value = self.main_window.open_file_dialog('Select a Image File', initial_directory=os.path.expanduser('~/'), file_types=['jpg', 'png'])
            self.file_name = self.file_name_input.value
            self.tmp_filen = f'{self.file_name}_temp.jpg'
            self.image_view.image = toga.Image(self.file_name)
            im = Image.open(self.file_name)
            self.width, self.height = im.size


        if widget.id == "font" or (widget.id == "add" and self.font_file_name == ""):
            self.font_label_loc.text = self.main_window.open_file_dialog('Select a ttf File', file_types=['ttf'])
            self.font_file_name = self.font_label_loc.text
            self.font_label_loc.text = "Font File : "+self.font_label_loc.text
        
        if widget.id == "add":
            if self.tmp_filen != "" and self.font_file_name != "":
                box = toga.Box(id=f'bnd_box_{self.bnd_box}')
                label = toga.Label('Text:', style=Pack(padding=5))
                self.text_list.append(toga.TextInput(id=f'lin_{self.bnd_box}', style=(Pack(padding=5))))
                x = toga.Label('x :', id=f'x_{self.bnd_box}', style=Pack(padding=5, color="red"))
                y = toga.Label('y :', id=f'y_{self.bnd_box}', style=Pack(padding=5, color="red"))
                
                x_in = toga.Slider(id=f'xin_{self.bnd_box}', style=Pack(padding=5), default=10, range=(0, self.width), on_slide=self.teleport)
                y_in = toga.Slider(id=f'yin_{self.bnd_box}', style=Pack(padding=5), default=10, range=(0, self.height), on_slide=self.teleport)
                
                box.add(label, self.text_list[self.bnd_box], x, x_in, y, y_in)
                self.coord.update({self.bnd_box:[10, 10, 36]})
                self.name_box.add(box)
                self.bnd_box += 1
                self.name_scroller.content = self.name_box
            else:
                self.fetch(widget)
        
        if widget.id == "open_csv":
            self.csv_file_name = toga.sources.ValueSource(value=self.main_window.open_file_dialog(title="Select the csv file", file_types=["csv"]))
            self.get_right_panel()
            # self.main_panel.content = [self.left_panel, self.get_right_panel()]
            # self.main_panel.refresh()
            # self.main_window.content = self.main_panel
        
        
            
    
    def teleport(self, widget, row=None):

        if len(self.file_name) != 0 and len(self.font_file_name):
            key = int(widget.id.split('_')[-1]) if '_' in widget.id else 0
            value = int(widget.value) if row == None else row.__dict__
            boxes = self.name_box.children
            change = False
            
            x_label = None
            y_label = None

            for box in boxes:
                ctrls = box.children
                for label in ctrls:
                    if label.id == f'x_{key}':
                        x_label = label
                    elif label.id == f'y_{key}':
                        y_label = label

            if widget.id.startswith('xin') and self.coord[key][0] != value:
                if x_label != None:  
                    x_label.text = f'x : {value}'
                self.coord[key][0] = value
                change = True

            elif widget.id.startswith('yin') and self.coord[key][1] != value:
                if y_label != None:
                    y_label.text = f'y : {value}'
                    print(y_label.text)                
                self.coord[key][1] = value
                change = True
            
            elif widget.id.startswith('fin'):
                self.f.text = f'Font Size  : {value} pts'
                
                for key in self.coord:
                    if self.coord[key][2] != value:
                        self.coord[key][2] = value
                        change = True
            
            elif widget.id == "sheet":
                # print(row.__dict__)
                for l_index, place_holder in enumerate(self.text_list):
                    for d_index, dict_key in enumerate(value['_attrs']):
                        if l_index == d_index-1:
                            place_holder.value = value[dict_key]
                            change = True
                            key = l_index
            
            
            print(x_label.text, y_label.text)
            x_label.refresh()
            y_label.refresh()
            self.name_box.refresh()
            self.global_ctrl_box.refresh()
            self.update_canvas(change, key)


    def update_canvas(self, change, key):
        if change and len(self.text_list[key].value) > 0:
            
            with open(self.font_file_name, 'rb') as FP:
                font_face =  ImageFont.truetype(FP, self.coord[0][-1])
                img = Image.open(self.file_name)
                new_image = ImageDraw.Draw(img)
                for key in self.coord:
                    new_image.text(self.coord[key][:-1], self.text_list[key].value, font=font_face)
                del font_face 
            img.save(self.tmp_filen)
            self.image_view.image = toga.Image(self.tmp_filen)
                
    def get_right_panel(self):
        right_panel = toga.Box(style=Pack(flex=1, direction=COLUMN))
        if self.csv_file_name.value is not None:
            with open(self.csv_file_name.value, 'r') as file_object:
                reader = csv.reader(file_object)
                temp = [row for row in reader]
                headings = ['_'+label.replace(' ', '_') if ' ' in label else '_'+label for label in temp[0]]
                data = temp[1:]
                # csv_widget = toga.Table(id="sheet", headings=headings, data=data, style=Pack(flex=1, padding=5), on_select=self.teleport)
                for heading in headings:
                    self.csv_widget.add_column(heading)
                self.csv_widget.data = data
                self.right_panel.refresh()
                
        # else:
        #     csv_widget = toga.Table(id="sheet", headings=['S.No','Name', 'Profession'], style=Pack(flex=1, padding=5), on_select=self.teleport)
        # open_csv = toga.Button("Open CSV File", id="open_csv", on_press=self.fetch)
        # right_panel.add(csv_widget)
        # right_panel.add(open_csv)
        # return right_panel

        

def plot(canv, color, col, row):
    with canv.fill(color=color) as dot:
        dot.rect(x=col, y=row, width=1, height=1)

def main():
    return CertificAtor()
