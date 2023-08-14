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
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_surfaceToCsv'
CMD_NAME = 'Surface to Points'
CMD_Description = 'Exports a surface as coodinates in a CSV file, based on a mesh.'

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

    # TODO Define the dialog for your command by adding different inputs to the command.
    # Create a Selection Command Input for surface
    surface_input = inputs.addSelectionInput('surface_input', 'Select surface', 'Select a surface body.')
    surface_input.setSelectionLimits(1, 1)
    surface_input.addSelectionFilter("SurfaceBodies")

    # Create a Dropdown Command Input for quality level
    quality_input = inputs.addDropDownCommandInput('quality_input', 'Mesh quality level',
                                                   adsk.core.DropDownStyles.TextListDropDownStyle)
    quality_input.listItems.add('Low', False, '')
    quality_input.listItems.add('Normal', False, '')
    quality_input.listItems.add('High', False, '')
    quality_input.listItems.add('Very High', True, '')

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
        mesh_bodies = root_comp.meshBodies
        mesh_quality_options = [adsk.fusion.TriangleMeshQualityOptions.LowQualityTriangleMesh,
                                adsk.fusion.TriangleMeshQualityOptions.NormalQualityTriangleMesh,
                                adsk.fusion.TriangleMeshQualityOptions.HighQualityTriangleMesh,
                                adsk.fusion.TriangleMeshQualityOptions.VeryHighQualityTriangleMesh]
        debug_info = ''

        # Get a reference to your command's inputs.
        inputs = args.command.commandInputs
        surface_input: adsk.core.SelectionCommandInput = inputs.itemById('surface_input')
        surface: adsk.fusion.BRepBody = surface_input.selection(0).entity
        quality_input: adsk.core.DropDownCommandInput = inputs.itemById('quality_input')
        quality_level = mesh_quality_options[quality_input.selectedItem.index]

        # Get control surface faces
        faces = []
        for i in range(surface.faces.count):
            faces.append(surface.faces.item(i))
        debug_info += f'The selected surface has: {surface.faces.count} faces.'

        # Create Mesh
        mesh_calc = surface.meshManager.createMeshCalculator()
        mesh_calc.setQuality(quality_level)
        t_mesh = mesh_calc.calculate()
        t_mesh_coordinates = t_mesh.nodeCoordinatesAsDouble
        t_mesh_indices = t_mesh.nodeIndices
        t_mesh_vector = t_mesh.normalVectorsAsDouble
        surface_mesh = mesh_bodies.addByTriangleMeshData(t_mesh_coordinates, t_mesh_indices, t_mesh_vector, [])
        debug_info += f'<br><br>Mesh generated with {quality_input.selectedItem.name} quality.'

        # Export points in csv file
        folder_path = filedialog.askdirectory()
        t_mesh_point = t_mesh.nodeCoordinates
        with open(folder_path + '/surface_points.csv', 'w', newline='') as file:
            writer = csv.writer(file)
            for p in t_mesh_point:
                x = units_manager.convert(p.x, 'cm', units_manager.defaultLengthUnits)
                y = units_manager.convert(p.y, 'cm', units_manager.defaultLengthUnits)
                z = units_manager.convert(p.z, 'cm', units_manager.defaultLengthUnits)
                writer.writerow([x, y, z])
        debug_info += f'<br><br>Exported CSV file: "{file.name}".'

        # Show Message
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
