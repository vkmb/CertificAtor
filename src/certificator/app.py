"""
Application for writing e-cetrificates
"""
import os
import csv
import time
import toga
import smtplib
from toga.style import Pack
from threading import Thread
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from toga.style.pack import COLUMN, ROW
from PIL import Image, ImageDraw, ImageFont
from email.mime.multipart import MIMEMultipart


class Mailer:
    def __init__(self, usr, profile, psk, server="smtp.gmail.com", port=587):
        try:
            self.server = str(server)
            self.port = int(port)
            self.mailer = smtplib.SMTP(self.server, self.port)
            self.usr = str(usr)
            self.profile = str(profile)
            self.mailer.ehlo()
            self.mailer.starttls()
            self.mailer.ehlo()
            self.mailer.login(user=self.usr, password=str(psk))
        except:
            self.mailer = None

    def send_mail(self, reciever_maildid, subject, message, image_fp=None):
        try:
            msg = MIMEMultipart()
            if image_fp is not None:
                img_data = open(image_fp, "rb").read()
                image = MIMEImage(img_data, name=os.path.basename(image_fp))
                msg.attach(image)
            msg["Subject"] = str(subject) if len(str(subject)) > 0 else "Test Mail"
            msg["From"] = f"{self.profile}<{self.usr}>"
            msg["To"] = ", ".join([reciever_maildid])
            text = MIMEText(str(message) if len(str(message)) > 0 else "Test Mail")
            msg.attach(text)
            self.mailer.sendmail(self.usr, reciever_maildid, msg.as_string())
            return ["mail has been initiated", time.ctime()]

        except:
            return ["mail failed to send", time.ctime()]


class CertificAtor(toga.App):
    def startup(self):
        """
        Construct and show the Toga application.

        Usually, you would add your application to a main content box.
        We then create a main window (with a name matching the app), and
        show the main window.
        """
        self.content = {
            "st": "Small Text",
            "mt": "Medium Text",
            "lt": "Large Text",
            "x": "X Coordinate",
            "y": "Y Coordinate",
            "th": "Table Column",
            "fs": "Font Size",
            "tl": "Character Length",
        }
        self.headings, self.data, self.write_certificate = None, None, None
        (
            self.image_file_name,
            self.tmp_filen,
            self.csv_file_name,
            self.font_file_name,
        ) = ("", "", "", "")
        self.image_file_width, self.image_file_height = 1000, 1000
        self.coord = {}
        self.bnd_box = 0
        self.text_list = []
        self.box_list = []
        self.x_in = []
        self.y_in = []
        self.left_panel = toga.Box(style=Pack(direction=COLUMN))

        self.file_name_label = toga.Label(
            "Certificate Image Name", style=Pack(padding=5)
        )
        self.font_label_loc = toga.Label("Font File : None", style=Pack(padding=5))
        self.fetch_button = toga.Button(
            "Open", id="open", style=Pack(padding=5), on_press=self.fetch
        )
        self.certificate_box = toga.Box(
            style=Pack(direction=ROW, padding=5),
            children=[self.file_name_label, self.fetch_button],
        )

        self.name_container = toga.Box(style=Pack(direction=ROW, padding=5))
        self.global_ctrl_box = toga.Box(style=Pack(direction=COLUMN, padding=5))
        self.add_name_button = toga.Button(
            "Add Label  ", id="add", style=Pack(padding=5, flex=1), on_press=self.fetch
        )
        self.add_font_file = toga.Button(
            "Change Font", id="font", style=Pack(padding=5, flex=1), on_press=self.fetch
        )
        self.global_ctrl_box.add(self.add_name_button, self.add_font_file)
        self.name_container.add(self.global_ctrl_box)

        self.name_box = toga.Box(style=Pack(direction=ROW, padding=5))
        self.name_scroller = toga.ScrollContainer(
            style=Pack(flex=1), content=self.name_box
        )
        self.image_view = toga.ImageView(style=Pack(flex=1, padding=5))
        self.image_tool = toga.SplitContainer(
            id="image_tool",
            style=Pack(flex=1, padding=5),
            content=[self.image_view, self.name_scroller],
        )
        self.send_emails = toga.Button(
            "Send Email",
            id="emails",
            style=Pack(padding=5, flex=1),
            enabled=True,
            on_press=self.get_mail_window,
        )
        self.open_csv = toga.Button(
            "Select a CSV File",
            id="open_csv",
            style=Pack(padding=5, flex=1),
            enabled=False,
            on_press=self.fetch,
        )
        self.write_certificate = toga.Button(
            "Write Certificate",
            id="",
            style=Pack(padding=5, flex=1),
            enabled=False,
            on_press=self.write_interface,
        )
        self.bottom_button_box = toga.Box(
            style=Pack(padding=5, direction=ROW),
            children=[self.open_csv, self.write_certificate, self.send_emails],
        )
        self.loading = toga.ProgressBar(style=Pack(padding=5))
        self.left_panel.add(
            self.image_tool,
            self.loading,
            self.name_container,
            self.certificate_box,
            self.font_label_loc,
            self.bottom_button_box,
        )
        self.main_window = toga.MainWindow(title=self.formal_name)
        self.csv_window = toga.Window(title=self.formal_name, closeable=False)
        self.main_window.content = self.left_panel
        self.csv_window.content = toga.Box(style=Pack(direction=COLUMN, padding=5))
        self.csv_window.show()
        self.main_window.show()

    def fetch(self, widget):
        if widget.id == "open" or (widget.id == "add" and self.image_file_name == ""):
            temp = self.main_window.open_file_dialog(
                "Select a Image File", file_types=["jpg", "png"]
            )
            self.image_file_name = (
                str(temp) if temp is not None else self.image_file_name
            )
            self.file_name_label.text = "Image File Location : " + self.image_file_name
            if self.image_file_name != "":
                self.tmp_filen = f"{self.image_file_name}_temp.jpg"
                self.image_view.image = toga.Image(self.image_file_name)
                im = Image.open(self.image_file_name)
                self.image_file_width, self.image_file_height = im.size
                del im

        if widget.id == "font" or (widget.id == "add" and self.font_file_name == ""):
            temp = self.main_window.open_file_dialog(
                "Select a ttf File", file_types=["ttf"]
            )
            self.font_file_name = str(temp) if temp is not None else self.font_file_name
            self.font_label_loc.text = "Font File : " + self.font_file_name

        if widget.id == "open_csv" or (widget.id == "add" and self.csv_file_name == ""):
            temp = self.main_window.open_file_dialog(
                title="Select the csv file", file_types=["csv"]
            )
            self.csv_file_name = str(temp) if temp is not None else self.csv_file_name
            self.get_right_panel()

        if widget.id == "add":
            if self.tmp_filen != "" and self.font_file_name != "":
                if not self.open_csv.enabled:
                    self.open_csv.enabled = True

                box = toga.Box(
                    id=f"bnd_box_{self.bnd_box}",
                    style=Pack(padding=5, direction=COLUMN),
                )
                label = toga.Label(
                    "Text:", id=f"textlabel_{self.bnd_box}", style=Pack(padding=5)
                )
                items = (
                    [""] + list(self.headings.keys())
                    if self.headings is not None
                    else []
                )
                self.text_list.append(
                    toga.TextInput(id=f"lin_{self.bnd_box}", style=(Pack(padding=5)))
                )
                table_label = toga.Selection(
                    id=f"tbl_{self.bnd_box}",
                    style=Pack(padding=5),
                    enabled=False if len(items) == 0 else True,
                    items=items,
                    on_select=self.teleport,
                )

                stl = toga.Label(
                    "Small Text Length : 0 chars",
                    id=f"stl_{self.bnd_box}",
                    style=Pack(padding=(5, 1, 5, 10), width=190),
                )
                stl_in = toga.Slider(
                    id=f"stlin_{self.bnd_box}",
                    style=Pack(padding=(5, 10, 5, 1)),
                    default=0,
                    range=(0, 255),
                    on_slide=self.teleport,
                )
                sts = toga.Label(
                    "Small Font Size : 0 pts",
                    id=f"sts_{self.bnd_box}",
                    style=Pack(padding=(5, 1, 5, 10), width=150),
                )
                sts_in = toga.Slider(
                    id=f"stsin_{self.bnd_box}",
                    style=Pack(padding=(5, 10, 5, 1)),
                    default=0,
                    range=(0, 100),
                    on_slide=self.teleport,
                )
                stx = toga.Label(
                    "x : 0", id=f"stx_{self.bnd_box}", style=Pack(padding=5, width=45)
                )
                sty = toga.Label(
                    "y : 0", id=f"sty_{self.bnd_box}", style=Pack(padding=5, width=45)
                )
                stx_in = toga.Slider(
                    id=f"stxin_{self.bnd_box}",
                    style=Pack(padding=5),
                    default=10,
                    range=(0, self.image_file_width),
                    on_slide=self.teleport,
                )
                sty_in = toga.Slider(
                    id=f"styin_{self.bnd_box}",
                    style=Pack(padding=5),
                    default=10,
                    range=(0, self.image_file_height),
                    on_slide=self.teleport,
                )
                stborder_line = toga.Divider(
                    id=f"bhst_{self.bnd_box}",
                    style=Pack(padding=5, height=2, width=100, flex=1),
                )

                mtl = toga.Label(
                    "Medium Text Length : 0 chars",
                    id=f"mtl_{self.bnd_box}",
                    style=Pack(padding=(5, 1, 5, 10), width=200),
                )
                mtl_in = toga.Slider(
                    id=f"mtlin_{self.bnd_box}",
                    style=Pack(padding=(5, 10, 5, 1)),
                    default=0,
                    range=(0, 255),
                    on_slide=self.teleport,
                )
                mts = toga.Label(
                    "Medium Font Size : 0 pts",
                    id=f"mts_{self.bnd_box}",
                    style=Pack(padding=(5, 1, 5, 10), width=160),
                )
                mts_in = toga.Slider(
                    id=f"mtsin_{self.bnd_box}",
                    style=Pack(padding=(5, 10, 5, 1)),
                    default=0,
                    range=(0, 100),
                    on_slide=self.teleport,
                )
                mtx = toga.Label(
                    "x : 0", id=f"mtx_{self.bnd_box}", style=Pack(padding=5, width=45)
                )
                mty = toga.Label(
                    "y : 0", id=f"mty_{self.bnd_box}", style=Pack(padding=5, width=45)
                )
                mtx_in = toga.Slider(
                    id=f"mtxin_{self.bnd_box}",
                    style=Pack(padding=5),
                    default=10,
                    range=(0, self.image_file_width),
                    on_slide=self.teleport,
                )
                mty_in = toga.Slider(
                    id=f"mtyin_{self.bnd_box}",
                    style=Pack(padding=5),
                    default=10,
                    range=(0, self.image_file_height),
                    on_slide=self.teleport,
                )
                mtborder_line = toga.Divider(
                    id=f"bhmt_{self.bnd_box}", style=Pack(padding=5, height=5, flex=1)
                )

                ltl = toga.Label(
                    "Large Text Length : 0 chars",
                    id=f"ltl_{self.bnd_box}",
                    style=Pack(padding=(5, 1, 5, 10), width=190),
                )
                ltl_in = toga.Slider(
                    id=f"ltlin_{self.bnd_box}",
                    style=Pack(padding=(5, 10, 5, 1)),
                    default=0,
                    range=(0, 255),
                    on_slide=self.teleport,
                )
                lts = toga.Label(
                    "Large Font Size : 0 pts",
                    id=f"lts_{self.bnd_box}",
                    style=Pack(padding=(5, 1, 5, 10), width=150),
                )
                lts_in = toga.Slider(
                    id=f"ltsin_{self.bnd_box}",
                    style=Pack(padding=(5, 10, 5, 1)),
                    default=0,
                    range=(0, 100),
                    on_slide=self.teleport,
                )
                ltx = toga.Label(
                    "x : 0", id=f"ltx_{self.bnd_box}", style=Pack(padding=5, width=45)
                )
                lty = toga.Label(
                    "y : 0", id=f"lty_{self.bnd_box}", style=Pack(padding=5, width=45)
                )
                ltx_in = toga.Slider(
                    id=f"ltxin_{self.bnd_box}",
                    style=Pack(padding=5),
                    default=10,
                    range=(0, self.image_file_width),
                    on_slide=self.teleport,
                )
                lty_in = toga.Slider(
                    id=f"ltyin_{self.bnd_box}",
                    style=Pack(padding=5),
                    default=10,
                    range=(0, self.image_file_height),
                    on_slide=self.teleport,
                )
                ltborder_line = toga.Divider(
                    id=f"bhlt_{self.bnd_box}", style=Pack(padding=5, height=5, flex=1)
                )

                box.add(
                    label,
                    self.text_list[self.bnd_box],
                    table_label,
                    stl,
                    stl_in,
                    sts,
                    sts_in,
                    stx,
                    stx_in,
                    sty,
                    sty_in,
                    stborder_line,
                    mtl,
                    mtl_in,
                    mts,
                    mts_in,
                    mtx,
                    mtx_in,
                    mty,
                    mty_in,
                    mtborder_line,
                    ltl,
                    ltl_in,
                    lts,
                    lts_in,
                    ltx,
                    ltx_in,
                    lty,
                    lty_in,
                    ltborder_line,
                )
                box_border_line = toga.Divider(
                    id=f"bv_{self.bnd_box}", direction=toga.Divider.VERTICAL
                )
                self.coord.update(
                    {
                        self.bnd_box: {
                            "st": {"x": 0, "y": 0, "fs": 0, "tl": 0},
                            "mt": {"x": 0, "y": 0, "fs": 0, "tl": 0},
                            "lt": {"x": 0, "y": 0, "fs": 0, "tl": 0},
                            "th": "",
                        }
                    }
                )
                self.name_box.add(box, box_border_line)
                self.bnd_box += 1
                self.name_scroller.content = self.name_box
            else:
                self.fetch(widget)

    def teleport(self, widget, row=None):

        if len(self.image_file_name) != 0 and len(self.font_file_name):
            key = int(widget.id.split("_")[-1]) if "_" in widget.id else 0
            value = (
                int(widget.value)
                if (row == None and "." in str(widget.value))
                else row.__dict__
                if row != None
                else widget.value
            )
            boxes = self.name_box.children
            change = False
            size = None
            (
                stx_label,
                sty_label,
                mtx_label,
                mty_label,
                ltx_label,
                lty_label,
                stl_label,
                mtl_label,
                ltl_label,
                sts_label,
                mts_label,
                lts_label,
            ) = (None, None, None, None, None, None, None, None, None, None, None, None)

            for box in boxes:
                ctrls = box.children
                for _label in ctrls:
                    if _label.id == f"stx_{key}":
                        stx_label = _label
                    elif _label.id == f"sty_{key}":
                        sty_label = _label
                    elif _label.id == f"mtx_{key}":
                        mtx_label = _label
                    elif _label.id == f"mty_{key}":
                        mty_label = _label
                    elif _label.id == f"ltx_{key}":
                        ltx_label = _label
                    elif _label.id == f"lty_{key}":
                        lty_label = _label
                    elif _label.id == f"stl_{key}":
                        stl_label = _label
                    elif _label.id == f"mtl_{key}":
                        mtl_label = _label
                    elif _label.id == f"ltl_{key}":
                        ltl_label = _label
                    elif _label.id == f"sts_{key}":
                        sts_label = _label
                    elif _label.id == f"mts_{key}":
                        mts_label = _label
                    elif _label.id == f"lts_{key}":
                        lts_label = _label

            if widget.id.startswith("stxin") and self.coord[key]["st"]["x"] != value:
                if stx_label != None:
                    stx_label.text = f"x : {value}"
                self.coord[key]["st"]["x"] = value
                size = "st"
                change = True

            elif widget.id.startswith("styin") and self.coord[key]["st"]["y"] != value:
                if sty_label != None:
                    sty_label.text = f"y : {value}"
                self.coord[key]["st"]["y"] = value
                size = "st"
                change = True

            elif widget.id.startswith("mtxin") and self.coord[key]["mt"]["x"] != value:
                if mtx_label != None:
                    mtx_label.text = f"x : {value}"
                self.coord[key]["mt"]["x"] = value
                size = "mt"
                change = True

            elif widget.id.startswith("mtyin") and self.coord[key]["mt"]["y"] != value:
                if mty_label != None:
                    mty_label.text = f"y : {value}"
                self.coord[key]["mt"]["y"] = value
                size = "mt"
                change = True

            elif widget.id.startswith("ltxin") and self.coord[key]["lt"]["x"] != value:
                if ltx_label != None:
                    ltx_label.text = f"x : {value}"
                self.coord[key]["lt"]["x"] = value
                size = "lt"
                change = True

            elif widget.id.startswith("ltyin") and self.coord[key]["lt"]["y"] != value:
                if lty_label != None:
                    lty_label.text = f"y : {value}"
                self.coord[key]["lt"]["y"] = value
                size = "lt"
                change = True

            elif widget.id.startswith("stlin") and self.coord[key]["st"]["tl"] != value:
                if stl_label != None:
                    stl_label.text = f"Small Text Length : {value} chars"
                self.coord[key]["st"]["tl"] = value

            elif widget.id.startswith("mtlin") and self.coord[key]["mt"]["tl"] != value:
                if mtl_label != None:
                    mtl_label.text = f"Medium Text Length : {value} chars"
                self.coord[key]["mt"]["tl"] = value

            elif widget.id.startswith("ltlin") and self.coord[key]["lt"]["tl"] != value:
                if ltl_label != None:
                    ltl_label.text = f"Long Text Length : {value} chars"
                self.coord[key]["lt"]["tl"] = value

            elif widget.id.startswith("stsin") and self.coord[key]["st"]["fs"] != value:
                if sts_label != None:
                    sts_label.text = f"Small Font Size : {value} pts"
                self.coord[key]["st"]["fs"] = value
                size = "st"
                change = True

            elif widget.id.startswith("mtsin") and self.coord[key]["mt"]["fs"] != value:
                if mts_label != None:
                    mts_label.text = f"Medium Font Size : {value} pts"
                self.coord[key]["mt"]["fs"] = value
                size = "mt"
                change = True

            elif widget.id.startswith("ltsin") and self.coord[key]["lt"]["fs"] != value:
                if lts_label != None:
                    lts_label.text = f"Long Font Size : {value} pts"
                self.coord[key]["lt"]["fs"] = value
                size = "lt"
                change = True

            elif widget.id.startswith("tbl") and self.coord[key]["th"] != value:
                self.coord[key]["th"] = str(value)

            elif widget.id == "sheet":
                for place_holder, l_index in zip(self.text_list, self.coord):
                    if self.coord[l_index]["th"] != "":
                        place_holder.value = value[
                            self.headings[self.coord[l_index]["th"]]
                        ]
                        change = True
                        key = l_index
            self.name_box.refresh()
            self.global_ctrl_box.refresh()
            self.update_canvas(change, key, size)

    def update_canvas(self, change, key, size=None):
        if change and len(self.text_list[key].value) > 0:
            img = Image.open(self.image_file_name)
            new_image = ImageDraw.Draw(img)
            for _key in self.coord:
                if _key == key and size is not None:
                    font_face = ImageFont.truetype(
                        self.font_file_name, self.coord[_key][size]["fs"]
                    )
                    new_image.text(
                        [self.coord[_key][size]["x"], self.coord[_key][size]["y"]],
                        self.text_list[_key].value,
                        font=font_face,
                    )
                else:
                    font_face = ImageFont.truetype(
                        self.font_file_name, self.coord[_key]["mt"]["fs"]
                    )
                    new_image.text(
                        [self.coord[_key]["mt"]["x"], self.coord[_key]["mt"]["y"]],
                        self.text_list[_key].value,
                        font=font_face,
                    )
                del font_face
            img.save(self.tmp_filen)
            self.image_view.image = toga.Image(self.tmp_filen)

    def get_right_panel(self):
        right_panel = toga.Box(style=Pack(flex=1, direction=COLUMN))
        if self.csv_file_name != "":
            with open(self.csv_file_name, "r") as file_object:
                reader = csv.reader(file_object)
                temp = [row for row in reader]
                heading_present = self.main_window.question_dialog(
                    "CSV File Configuration", "Does the csv file has a heading column ?"
                )
                if not heading_present:
                    self.headings = {
                        f"Column {_index}": f"column_{_index}"
                        for _index, label in enumerate(temp[0], 1)
                    }
                    self.data = temp[0:]
                else:
                    self.headings = {
                        label: ("_" + label.replace(" ", "_")).lower()
                        if " " in label
                        else ("_" + label).lower()
                        for label in temp[0]
                    }
                    self.data = temp[1:]
                csv_widget = toga.Table(
                    id="sheet",
                    headings=list(self.headings.keys()),
                    data=self.data,
                    style=Pack(flex=1, padding=5),
                    on_select=self.teleport,
                    accessors=list(self.headings.values()),
                )
                for box in self.name_box.children:
                    ctrls = box.children
                    for _label in ctrls:
                        if _label.id.startswith("tbl_"):
                            _label.items = [""] + list(self.headings.keys())
                            _label.enabled = True
                        if _label.id.startswith("lin_"):
                            _label.value = ""
                self.write_certificate.enabled = True
        else:
            csv_widget = toga.Table(
                id="sheet", headings=[], style=Pack(flex=1, padding=5)
            )
        right_panel.add(csv_widget)
        self.csv_window.content = right_panel
        self.csv_window.show()

    def write_interface(self, widget):
        message = "Writing Options"
        self.consolidated = {}
        # confirm the label headers
        for _index in self.coord:
            consolidated = {}
            message += f"\nLabel {_index+1} "
            for _key in self.coord[_index]:
                if _key != "th" and self.coord[_index][_key]["fs"] == 0:
                    continue
                if _key == "th" and (
                    self.coord[_index]["st"]["fs"] != 0
                    or self.coord[_index]["mt"]["fs"] != 0
                    or self.coord[_index]["lt"]["fs"] != 0
                ):
                    consolidated.update(
                        {
                            _key: list(self.headings.keys()).index(
                                self.coord[_index][_key]
                            )
                        }
                    )
                    message += (
                        f"\n\t-> {self.content[_key]} : {self.coord[_index][_key]}"
                    )
                    continue
                else:
                    consolidated.update({_key: self.coord[_index][_key]})
                    message += f"\n\t-> Category : {self.content[_key]}"
                    for _sub_key in self.coord[_index][_key]:
                        message += f"\n\t\t-> {self.content[_sub_key]} : {self.coord[_index][_key][_sub_key]}"
            message += "\n"
            self.consolidated.update({_index: consolidated})

        go_ahead = self.main_window.confirm_dialog(
            title="Write Conform", message=message
        )

        if go_ahead:
            self.destination_folder = self.main_window.select_folder_dialog(
                title="Select destination folder"
            )
            self.destination_folder = self.destination_folder[0]
            # self.thread = Thread(target=self.write_certificates, daemon=True)
            # self.thread.start()
            self.write_certificates()
            if self.csv_log_file_name != "":
                self.get_mail_window(None)

    def write_certificates(self):
        to_write = {}
        self.csv_log_file_name = ""
        logs = []
        flag = False
        self.loading.max = len(self.data)
        self.loading.start()

        for sno, row in enumerate(self.data, 1):
            to_write = {
                _label: row[self.consolidated[_label]["th"]]
                for _label in self.consolidated
            }
            image = Image.open(self.image_file_name)
            new_image = ImageDraw.Draw(image)
            for label in to_write:
                font_face = ""
                # build cases
                if (
                    "st" in self.consolidated[label].keys()
                    and len(to_write[label]) <= self.consolidated[label]["st"]["tl"]
                ):
                    font_face = ImageFont.truetype(
                        self.font_file_name, self.consolidated[label]["st"]["fs"]
                    )
                    new_image.text(
                        [
                            self.consolidated[label]["st"]["x"],
                            self.consolidated[label]["st"]["y"],
                        ],
                        to_write[label],
                        font=font_face,
                    )
                if (
                    "mt" in self.consolidated[label].keys()
                    and self.consolidated[label]["st"]["tl"]
                    < len(to_write[label])
                    <= self.consolidated[label]["mt"]["tl"]
                ):
                    font_face = ImageFont.truetype(
                        self.font_file_name, self.consolidated[label]["mt"]["fs"]
                    )
                    new_image.text(
                        [
                            self.consolidated[label]["mt"]["x"],
                            self.consolidated[label]["mt"]["y"],
                        ],
                        to_write[label],
                        font=font_face,
                    )
                if (
                    "lt" in self.consolidated[label].keys()
                    and self.consolidated[label]["mt"]["tl"]
                    < len(to_write[label])
                    <= self.consolidated[label]["lt"]["tl"]
                ):
                    font_face = ImageFont.truetype(
                        self.font_file_name, self.consolidated[label]["lt"]["fs"]
                    )
                    new_image.text(
                        [
                            self.consolidated[label]["lt"]["x"],
                            self.consolidated[label]["lt"]["y"],
                        ],
                        to_write[label],
                        font=font_face,
                    )
                del font_face
            fn = f'{sno}.{"_".join(list(to_write.values()))}.png'
            written_certificate_name = os.path.join(self.destination_folder, fn)
            image.save(written_certificate_name)
            temp = row + [os.path.abspath(written_certificate_name)]
            logs.append(temp)
            self.loading.value = sno

        # log to file
        if self.csv_file_name != "":
            self.csv_log_file_name = (
                str(
                    self.csv_file_name
                    if self.csv_file_name != ""
                    else f"temp_{time.ctime()}.csv"
                ).split(".csv")[0]
                + "_file_log.csv"
            )
            with open(self.csv_log_file_name, "w") as csv_log_fp:
                csv_logger = csv.writer(csv_log_fp)
                csv_logger.writerows(logs)

        self.loading.stop()
        self.loading.value = 0

    def get_mail_window(self, widget):
        if widget != None:
            temp = self.main_window.open_file_dialog(
                "Select a csv file with mail ids", file_types=["csv"]
            )
            if temp is not None:
                self.csv_log_file_name = str(temp)
            else:
                return
        go_ahead = self.main_window.question_dialog(
            title="SenD Emails",
            message=f"Do you want to use mail ids from following file \n{self.csv_log_file_name}",
        )

        if go_ahead:
            with open(self.csv_log_file_name, "r") as file_object:
                reader = csv.reader(file_object)
                temp = [row for row in reader]
                heading_present = self.main_window.question_dialog(
                    "CSV File Configuration", "Does the csv file has a heading column ?"
                )
                if not heading_present:
                    self.email_headings = {
                        f"Column {_index+1}": _index
                        for _index, label in enumerate(temp[0], 0)
                    }
                    self.email_data = temp
                else:
                    self.email_headings = {
                        label: _index for _index, label in enumerate(temp[0], 0)
                    }
                    self.email_data = temp[1:]

            self.mail_id_label = toga.Label(
                text="Mail Id      : ", id="mail_id", style=Pack(padding=10)
            )
            self.mail_id_usin = toga.TextInput(
                id="mailid_in", style=Pack(padding=10, flex=1)
            )
            self.mail_id_box = toga.Box(
                id="mail_box",
                style=Pack(flex=1, padding=5),
                children=[self.mail_id_label, self.mail_id_usin],
            )

            self.mail_pn_label = toga.Label(
                text="Profile Name : ", id="mail_pn", style=Pack(padding=10)
            )
            self.mail_pn_usin = toga.TextInput(
                id="mailpn_in", style=Pack(padding=10, flex=1)
            )
            self.mail_pn_box = toga.Box(
                id="mail_box",
                style=Pack(flex=1, padding=5),
                children=[self.mail_pn_label, self.mail_pn_usin],
            )

            self.pass_label = toga.Label(
                text="Password     : ", id="pass", style=Pack(padding=10)
            )
            self.pass_usin = toga.PasswordInput(
                id="pass_in", style=Pack(padding=10, flex=1)
            )
            self.pass_box = toga.Box(
                id="pass_box",
                style=Pack(flex=1, padding=5),
                children=[self.pass_label, self.pass_usin],
            )

            self.sub_label = toga.Label(
                text="Subject      : ", id="subject", style=Pack(padding=10)
            )
            self.sub_usin = toga.TextInput(id="sub_in", style=Pack(padding=10, flex=1))
            self.sub_box = toga.Box(
                id="sub_box",
                style=Pack(flex=1, padding=5),
                children=[self.sub_label, self.sub_usin],
            )

            self.image_column_index = toga.Selection(
                id="image_column_in",
                style=Pack(padding=10, flex=1),
                items=["Image Column Name"] + list(self.email_headings.keys()),
            )
            self.email_column_index = toga.Selection(
                id="email_column_in",
                style=Pack(padding=10, flex=1),
                items=["Email Column Name"] + list(self.email_headings.keys()),
            )
            self.mess_usin = toga.MultilineTextInput(
                id="mess_in", style=Pack(padding=10, flex=1), initial="Content of mail"
            )
            self.email_pb = toga.ProgressBar(
                id="email_pb", style=Pack(padding=10, flex=1)
            )
            self.sm_button = toga.Button(
                label="Send Mail",
                id="sendmail",
                style=Pack(padding=10, flex=1),
                on_press=self.send_mail,
            )
            self.st_button = toga.Button(
                label="Send Test Mail",
                id="sendtestmail",
                style=Pack(padding=10, flex=1),
                on_press=self.send_mail,
            )
            self.cm_button = toga.Button(
                label="Cancel",
                id="cancelmail",
                style=Pack(padding=10, flex=1),
                on_press=self.send_mail,
            )
            self.mess_box = toga.Box(
                id="mess_box",
                style=Pack(flex=1, padding=5, direction=COLUMN),
                children=[
                    self.image_column_index,
                    self.email_column_index,
                    self.mess_usin,
                    self.email_pb,
                    self.sm_button,
                    self.st_button,
                    self.cm_button,
                ],
            )

            self.mail_content_box = toga.Box(
                id="mail_content_box",
                style=Pack(direction=COLUMN, flex=1),
                children=[
                    self.mail_pn_box,
                    self.mail_id_box,
                    self.pass_box,
                    self.sub_box,
                    self.mess_box,
                ],
            )
            self.email_window = toga.Window(
                id="email_dialog", title="Email Configuration"
            )
            self.email_window.content = self.mail_content_box
            self.email_window.show()

    def send_mail(self, widget):

        if widget.id.startswith("sendm") or widget.id.startswith("sendt"):

            email_index, image_path_index = None, None

            if (
                not len(self.mail_id_usin.value) > 3
                or "@" not in str(self.mail_id_usin.value)
                or "." not in str(self.mail_id_usin.value)
            ):
                self.main_window.error_dialog("Mail Id Error", "Enter a valid mail")
                return
            if not len(self.mail_pn_usin.value) > 2:
                self.main_window.error_dialog(
                    "Profile Name Error", "Profile name not typed"
                )
                return
            if not len(self.pass_usin.value) > 5:
                self.main_window.error_dialog("Password Error", "Password not typed")
                return
            if not len(self.sub_usin.value) > 1:
                self.main_window.error_dialog("Subject Error", "Subject not typed")
                return
            if not len(self.sub_usin.value) > 2:
                self.main_window.error_dialog(
                    "Mail Content Error", "Mail content not typed"
                )
                return
            if str(self.email_column_index.value) == "Email Column Name":
                self.main_window.error_dialog(
                    "Email Column Name Error", "Email column not selected"
                )
                return
            else:
                email_index = self.email_headings[str(self.email_column_index.value)]

            if str(self.image_column_index.value) != "Image Column Name":
                image_path_index = self.email_headings[
                    str(self.image_column_index.value)
                ]

            self.loading.max = len(self.email_data)
            mailer = Mailer(
                usr=self.mail_id_usin.value,
                profile=self.mail_pn_usin.value,
                psk=self.pass_usin.value,
            )
            time.sleep(20)

            if mailer.mailer is not None:

                if widget.id.startswith("sendtest"):
                    image_path = (
                        None
                        if image_path_index is None
                        else self.email_data[0][image_path_index]
                    )
                    print(
                        self.mail_id_usin.value,
                        self.sub_usin.value,
                        self.mess_usin.value,
                        image_path,
                    )
                    feedback = mailer.send_mail(
                        self.mail_id_usin.value,
                        self.sub_usin.value,
                        self.mess_usin.value,
                        image_path,
                    )
                    self.main_window.info_dialog(
                        "Mail Log",
                        f"A {feedback[0]} to {self.mail_id_usin.value} at {feedback[1]}",
                    )
                    return

                elif widget.id.startswith("sendma"):
                    self.email_pb.max = len(self.email_data)
                    self.email_pb.start()
                    logs = []
                    for row in self.email_data:
                        print(row[email_index])
                        feedback = mailer.send_mail(
                            (row[email_index]).strip(),
                            self.sub_usin.value,
                            self.mess_usin.value,
                            None if image_path_index is None else row[image_path_index],
                        )
                        logs.append(row + feedback)
                        self.email_pb.value = self.email_data.index(row) + 1
                        time.sleep(1)
                    self.email_pb.stop()
                    self.email_pb.value = 0

                    mail_log = (
                        self.csv_log_file_name.split(".csv")[0]
                        + f"_maillogs_{time.ctime()}.csv"
                    )
                    with open(mail_log, "w") as file_object:
                        writer = csv.writer(file_object)
                        writer.writerows(logs)

                    self.main_window.info_dialog(
                        title="Mail Log",
                        message=f"The mail logs is available in {mail_log}",
                    )
                mailer.mailer.quit()
            else:
                self.main_window.error_dialog(
                    "Mail Service Error",
                    "Please check the mailid, password & internet connection",
                )
                return

        elif widget.id.startswith("cancelm"):
            self.email_window.close()


def main():
    return CertificAtor()
