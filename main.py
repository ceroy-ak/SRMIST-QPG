import PySimpleGUI as sg
import os
from question import questions


layout = [[sg.Text("Welcome to SRMIST Question Paper Generator")],
          [sg.Text("Source File"), sg.In(), sg.FileBrowse(file_types=(("Excel Files(*.xlsx only)", "*.xlsx"),))],
          [sg.Text("Destination"), sg.In(), sg.FolderBrowse()],
          [sg.InputCombo( ['Cycle Test 1','Cycle Test 2','Cycle Test 3','University Exam'],
                       default_value=None,
                       size=(None, None),
                       auto_size_text=True,
                       background_color=None,
                       text_color=None,
                       change_submits=False,
                       enable_events=False,
                       readonly=True,
                       disabled=False,
                       key=None ,
                       pad=None,
                       tooltip=None,
                       visible=True)],
          [sg.Text("Difficulty"),
            sg.Slider(range=(0,100),
                        default_value=25,
                        resolution=25,
                        orientation='h',
                        border_width=None,
                        relief=None,
                        change_submits=False,
                        disabled=False,
                        size=(None, None),
                        font=None,
                        background_color=None,
                        text_color=None,
                        key=None,
                        pad=None,
                        tooltip="Choose Difficulty",
                        visible=True)],
          [sg.Submit(), sg.Cancel()]
         ]


if __name__ == "__main__" :

    #Windows
    window = sg.Window("SRMIST QPG",
            default_element_size=sg.DEFAULT_ELEMENT_SIZE,
            default_button_element_size=(None,None),
            auto_size_text=True,
            auto_size_buttons=True,
            location=(None,None),
            size=(None,None),
            element_padding=None,
            button_color=None,
            font=None,
            progress_bar_color=(None,None),
            background_color=None,
            border_depth=None,
            auto_close=False,
            auto_close_duration=sg.DEFAULT_AUTOCLOSE_TIME,
            icon="srm.ico",
            force_toplevel=False,
            alpha_channel=1,
            return_keyboard_events=False,
            use_default_focus=True,
            text_justification=True,
            no_titlebar=False,
            grab_anywhere=False,
            keep_on_top=False,
            resizable=False,
            disable_close=False,
            disable_minimize=False,
            right_click_menu=None).Layout(layout)

    #Reading the events
    button, values = window.Read()
    #print(values)

    #Send the Values
    result = questions(values)

    if result is True:
        sg.Popup('The Questions are successfully generated')
        filename = str(values[1]) + '/Questions.pdf'
        os.system("start " + filename)
    else:
        sg.PopupError('The file could not be generated. Please Contact the administrator')

