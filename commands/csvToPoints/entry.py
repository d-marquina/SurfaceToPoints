import adsk.core
import adsk.fusion
import os
import tkinter
from tkinter import filedialog
import csv
import traceback
from ...lib import fusion360utils as futil
from ... import config

app = adsk.core.Application.get()
ui = app.userInterface
tkinter.Tk().withdraw()


# TODO *** Specify the command identity information. ***
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_csvToPoints'
CMD_NAME = 'CSV to Points'
CMD_Description = 'Imports points from a CSV file.'

# Specify that the command will be promoted to the panel.
IS_PROMOTED = True

# TODO *** Define the location where the command button will be created. ***
# This is done by specifying the workspace, the tab, and the panel, and the 
# command it will be inserted beside. Not providing the command to position it
# will insert it at the end.
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidScriptsAddinsPanel'
COMMAND_BESIDE_ID = 'ScriptsManagerCommand'

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# Local list of event handlers used to maintain a reference, so
# they are not released and garbage collected.
local_handlers = []


# Executed when add-in is run.
def start():
    # Create a command Definition.
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

    # Define an event handler for the command created event. It will be called when the button is clicked.
    futil.add_handler(cmd_def.commandCreated, command_created)

    # ******** Add a button into the UI so the user can run the command. ********
    # Get the target workspace the button will be created in.
    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Get the panel the button will be created in.
    panel = workspace.toolbarPanels.itemById(PANEL_ID)

    # Create the button command control in the UI after the specified existing command.
    control = panel.controls.addCommand(cmd_def, COMMAND_BESIDE_ID, False)

    # Specify if the command is promoted to the main toolbar. 
    control.isPromoted = IS_PROMOTED


# Executed when add-in is stopped.
def stop():
    # Get the various UI elements for this command
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    # Delete the button command control
    if command_control:
        command_control.deleteMe()

    # Delete the command definition
    if command_definition:
        command_definition.deleteMe()


# Function that is called when a user clicks the corresponding button in the UI.
# This defines the contents of the command dialog and connects to the command related events.
def command_created(args: adsk.core.CommandCreatedEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Created Event')

    # https://help.autodesk.com/view/fusion360/ENU/?contextId=CommandInputs
    inputs = args.command.commandInputs

    # TODO Connect to the events that are needed by this command.
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.executePreview, command_preview, local_handlers=local_handlers)
    futil.add_handler(args.command.validateInputs, command_validate_input, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)


# This event handler is called when the user clicks the OK button in the command dialog or 
# is immediately called after the created event not command inputs were created for the dialog.
def command_execute(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Execute Event')

    # TODO ******************************** Your code here ********************************
    try:
        # Get main objects
        product = app.activeProduct
        if not product:
            ui.messageBox('No active Fusion design', 'No design')
            return
        design = adsk.fusion.Design.cast(product)
        units_manager = design.unitsManager
        root_comp = design.rootComponent
        sketches = root_comp.sketches
        default_length_units = units_manager.defaultLengthUnits
        debug_info = ''

        # Get CSV file
        csv_filename = filedialog.askopenfilename()
        debug_info += f'CSV file name: {csv_filename}.'

        # Get CSV data
        csv_data = []
        row_count = 0
        with open(csv_filename) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            for row in csv_reader:
                point_data = [row[0], row[1], row[2]]
                csv_data.append(point_data)
                row_count += 1
        debug_info += f'<br>Number of points: {row_count}.'

        # Create sketches
        p_x_s = 100  # points per sketch
        n_sketches = row_count // 100
        rem_points = row_count % 100
        if rem_points >= 2:
            rem_sketch = 1
        else:
            rem_sketch = 0
        p_idx = 0  # points index
        debug_info += f'<br>Number of Sketches: {n_sketches + rem_sketch}'
        debug_info += f'<br>Remainder points: {rem_points}'
        for n_sketch in range(n_sketches + rem_sketch):
            csv_sketch = sketches.add(root_comp.xYConstructionPlane)
            if n_sketch is n_sketches:
                n_points = rem_points - 1
            else:
                n_points = p_x_s - 1
            for i in range(n_points):
                min_point_c = []
                max_point_c = []
                for a in range(3):
                    min_point_c.append(units_manager.convert(float(csv_data[i+p_idx][a]), default_length_units, 'cm'))
                for b in range(3):
                    max_point_c.append(units_manager.convert(float(csv_data[i+p_idx+1][b]), default_length_units, 'cm'))
                min_point = adsk.core.Point3D.create(min_point_c[0], min_point_c[1], min_point_c[2])
                max_point = adsk.core.Point3D.create(max_point_c[0], max_point_c[1], max_point_c[2])
                csv_sketch.sketchCurves.sketchLines.addByTwoPoints(min_point, max_point)
            p_idx += p_x_s

        # Show Debug Info
        ui.messageBox(debug_info)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# This event handler is called when the command needs to compute a new preview in the graphics window.
def command_preview(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Preview Event')
    inputs = args.command.commandInputs


# This event handler is called when the user changes anything in the command dialog
# allowing you to modify values of other inputs based on that change.
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    inputs = args.inputs

    # General logging for debug.
    futil.log(f'{CMD_NAME} Input Changed Event fired from a change to {changed_input.id}')


# This event handler is called when the user interacts with any of the inputs in the dialog
# which allows you to verify all the inputs are valid and enables the OK button.
def command_validate_input(args: adsk.core.ValidateInputsEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Validate Input Event')

    inputs = args.inputs
        

# This event handler is called when the command terminates.
def command_destroy(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Destroy Event')

    global local_handlers
    local_handlers = []
