# Diego Suarez
VERSION = 'beta'

# importing sys
import sys

# adding Folder_2/subfolder to the system path
sys.path.insert(0, r'C:\Users\dsuas\VV Ingenieros Dropbox\06 Research\2022 MACROS Y APIs\ETABS APIs DSS\GIT\pytabs\src')


import os
import pytabs
import xlwings as xw
from datetime import datetime

#  TODO: incluir opcion de backup, imprimir nombre de archivo conectado, agregar combinacion de cargas

# workbook mock caller
WB_FN = os.path.join(os.path.dirname(__file__), 'LoadCasesMaster.xlsm')

NOW = datetime.now()

# batch workbook configuration
WB_BATCH_IN_SHEET = 'Main'

WB_HEADER_IN_RANGE = {'job_number': 'B2',
                      'project': 'B3',
                      'element': 'B4',
                      'designer': 'B5',
                      'date': 'B6',
                      'backup': 'B7',
                      'model_fp':'B8',
                      'model_open': 'B9',
                      }
PRM_Modal = {}

WB_RUNS_IN_RANGE = 'A13:E13'

# WB_RUN_RESULTS_LOCATIONS = 'CT13:DY13'

WB_BATCH_RESULTS_OUT = 'A14'


def read_model_parameters(wb : xw.Book):
    
    MODEL_PARAMETERS = {'Espectro': 'E13:E18',
                          'Modal': 'E13:E20',
                          'SHELL_LP': 'F14:F200',
                          }
    

def read_batch_input(batch_wb : xw.Book):
    input_sheet = batch_wb.sheets[WB_BATCH_IN_SHEET]
    # read header information
    header_input = {}
    for key, range in WB_HEADER_IN_RANGE.items():
        header_input[key] = input_sheet.range(range).value
    
    batch_runs_input =  input_sheet.range(WB_RUNS_IN_RANGE).options(empty='', expand='down').value
    batch_runs_headers = batch_runs_input[0]
    batch_runs_data = batch_runs_input[1:]
    
    return header_input, batch_runs_headers, batch_runs_data 

def write_batch_results(batch_wb : xw.Book, run_results):
    input_sheet = batch_wb.sheets[WB_BATCH_IN_SHEET]
    input_sheet.range(WB_BATCH_RESULTS_OUT).expand('down').clear_contents()
    input_sheet.range(WB_BATCH_RESULTS_OUT).options(expand='down', transpose = True).value = run_results
    
    
def delete_load_patterns(etabs_model):
    #get existing load patterns and delete
    ExistLPs = etabs_model.load_patterns.get_name_list()  
    if ExistLPs[0]!='':
        [etabs_model.load_patterns.delete(name) for name in ExistLPs]    
        print(f"{len(ExistLPs)} load patterns were deleted: {ExistLPs[0:5]}...")
    

def delete_load_cases(etabs_model):
    #get existing loadcases and delete
    ExistLCs = etabs_model.load_cases.get_name_list()  
    if ExistLCs[0]!='':
        [etabs_model.load_cases.delete(name)   for name in ExistLCs if name != 'Modal' ]   
        [etabs_model.load_cases.delete('Modal')   for name in ExistLCs if name == 'Modal' ] 
        print(f"{len(ExistLCs)} load cases were deleted: {ExistLCs[0:5]} ...")
    
def assign_load_patterns(etabs_model,tipoLP):
    """ par LoadPatternsVV: Dictionary for load pattern names containing load types, and self self_weight_multipliers
        par tipo: specify type of load pattern to be applied (VVIng, BA)
    """
    LoadPatternsVV = {'Do': [etabs_model.eLoadPatternType.Dead,1,True],
                    'DD': [etabs_model.eLoadPatternType.SuperDead,0,True],
                    'Dcons': [etabs_model.eLoadPatternType.SuperDead,0,True],
                    'ELLx': [etabs_model.eLoadPatternType.Quake,0,False],
                    'Elly': [etabs_model.eLoadPatternType.Quake,0,False],
                    'ERSx': [etabs_model.eLoadPatternType.Quake,0,False],
                    'ERSy': [etabs_model.eLoadPatternType.Quake,0,False],
                    'ELLx_M': [etabs_model.eLoadPatternType.Quake,0,False],
                    'ELLy_M': [etabs_model.eLoadPatternType.Quake,0,False],
                    'ERSx_M': [etabs_model.eLoadPatternType.Quake,0,False],
                    'ERSy_M': [etabs_model.eLoadPatternType.Quake,0,False],
                    'H': [etabs_model.eLoadPatternType.Other,0,True],
                    'La': [etabs_model.eLoadPatternType.Live,0,True],
                    'Lb': [etabs_model.eLoadPatternType.Live,0,True],
                    'Lc': [etabs_model.eLoadPatternType.Live,0,True],
                    'Ld': [etabs_model.eLoadPatternType.Live,0,True],
                    'Lr': [etabs_model.eLoadPatternType.Rooflive,0,True],
                    'Nx': [etabs_model.eLoadPatternType.Notional,0,False],
                    'Ny': [etabs_model.eLoadPatternType.Notional,0,False],
                    'R': [etabs_model.eLoadPatternType.Other,0,True],
                    'W0': [etabs_model.eLoadPatternType.Wind,0,False],
                    'W90': [etabs_model.eLoadPatternType.Wind,0,False] 
                    }
    
    for i, name in enumerate(LoadPatternsVV.keys()):
        if tipoLP == 'VVIng':
            ret = etabs_model.load_patterns.add(name,LoadPatternsVV[name][0],
                                            LoadPatternsVV[name][1],False)
        
def add_CSCR2010_RS(etabs_model,RS_parameters=['ESPECTRO1','III','S2','GroupD','3','2','0.05']):
    
    # add response spectrum case
    table = etabs_model.database_tables.get_table_details('Functions - Response Spectrum - Costa Rica Seismic Code 2010')
    tableData = etabs_model.database_tables.get_table_fields('Functions - Response Spectrum - Costa Rica Seismic Code 2010')
    table_array = etabs_model.database_tables.get_table_data_array(table,edit_mode=True)
    new_table_data = RS_parameters
    etabs_model.database_tables.set_table_data_array_edit(table=table,
                                                          field_keys=table_array.fields_included,
                                                          table_data=new_table_data)
    db_edit_log = etabs_model.database_tables.apply_table_edits(True)
    print(db_edit_log['import_log'])
    etabs_model.database_tables.discard_table_edits()
    
# def add_Modal_Case(etabs_model,analysis_type):
    
#     match analysis_type:
#             case 'Ritz':
#                 table_name = 'Modal Case Definitions Ritz'
#             case 'Eigen':
#                 table_name = 'Modal Case Definitions Eigen'
#             case _:
#                 table_name = 0

    
#     table = etabs_model.database_tables.get_table_details('Modal Case Definitions Eigen')
#     tableData = etabs_model.database_tables.get_table_fields('Functions - Response Spectrum - Costa Rica Seismic Code 2010')
#     table_array = etabs_model.database_tables.get_table_data_array(table,edit_mode=True)
#     new_table_data = Modal_parameters
#     etabs_model.database_tables.set_table_data_array_edit(table=table,
#                                                           field_keys=table_array.fields_included,
#                                                           table_data=new_table_data)
#     db_edit_log = etabs_model.database_tables.apply_table_edits(True)
#     print(db_edit_log['import_log'])
#     etabs_model.database_tables.discard_table_edits()
        

def assign_load_combos(etabs_model,tipoAnalisis):
    
    """ par LoadPatternsVV: Dictionary for load pattern names containing load types, and self self_weight_multipliers
        par tipoAnalisis: specify type of load combinations to be set
        
    """
    
    LoadCombos = {
        '1': {'Do': 1, 'DD': 1},
        '2': {'La': 1,'Lb': 1,'Lc': 1,'Ld': 1,'Lr': 1},
        '3': {'Do': 1, 'DD': 1,'La': 1, 'Lb': 1, 'Lc': 1, 'Ld': 1, 'Lr': 1},
        }
    
def assign_modal_case(etabs_model):
    

    etabs_model.load_cases.load_cases.ModalEigen.SetInitialCase("LCASE1", "SN1")
    etabs_model.load_cases.load_cases.ModalEigen.SetParameters("LCASE1", "SN1")

    
    

def main():
    
    print('Initiating ...', end='')
    batch_wb = xw.Book.caller()
    header_input, batch_runs_headers, batch_runs_data = read_batch_input(batch_wb)
    print('done.')
    
    model_fp = header_input['model_fp']
    model_is_open = header_input['model_open']

    # Check if the model is open
    if model_is_open.lower() == 'yes':
        model_is_open = True
    else:
        model_is_open = False

    # if model open set no model path
    if model_is_open:
        model_fp = ''
        print('\nAttaching to open ETABS instance.')
    else:
        print(f"\nOpening ETABS model: {model_fp}.")
    
    # substantiate pyTABS EtabsModel
    etabs_model = pytabs.EtabsModel(attach_to_instance=model_is_open, model_path=model_fp, backup=False)
    
    # get model name
    model_file = etabs_model.get_filename()
    print('connected to file ',model_file)
    
    # set the model units to kgf-m-C
    etabs_model.set_present_units(etabs_model.eUnits.kgf_m_C)
    print(f"Units set to kgf_m_C")
    
    etabs_model.set_model_is_locked(False)
    
    # Delete existing LsC and LPs
    delete_load_cases(etabs_model)
    delete_load_patterns(etabs_model)

    #Create new load Patterns
    assign_load_patterns(etabs_model,'VVIng')
    print(f"VVIngenieria load patterns were added")
    
    
    add_CSCR2010_RS(etabs_model)
    
    
    # etabs_model.load_cases.load_cases.StaticLinear.SetCase('AA')
    
    
    print('done.', end='')
    print('\nRun complete, press any key to exit.')
    input('')

    # exit ETABS if not attached
    if not model_is_open:
        etabs_model.exit_application()



if __name__ == "__main__":
    xw.Book(WB_FN).set_mock_caller()
    main()