# pyTABS Example: GroupAssigner
# Sam Cubis
VERSION = '230202'

import os
import pytabs
import xlwings as xw
from datetime import datetime

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
                      'notes': 'B7',
                      'model_fp':'B8',
                      'model_open': 'B9'}

WB_RUNS_IN_RANGE = 'A13:E13'

# WB_RUN_RESULTS_LOCATIONS = 'CT13:DY13'

WB_BATCH_RESULTS_OUT = 'A14'


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
    
    
def assign_load_patterns_VV(ETABS_MODEL):
    
    """ par LoadPatternsVV: Dictinoary with load pattern names, load type, and self self_weight_multiplier
        
    
    
    """
    LoadPatternsVV = {'Do': [1,1],
                    'DD': [2,0],
                    'ELLx': [5,0],
                    'Elly': [5,0],
                    'ERSx': [5,0],
                    'ERSy': [5,0],
                    'ELLx_M': [5,0],
                    'ELLy_M': [5,0],
                    'ERSx_M': [5,0],
                    'ERSy_M': [5,0],
                    'H': [8,0],
                    'La': [3,0],
                    'Lb': [3,0],
                    'Lc': [3,0],
                    'Ld': [3,0],
                    'Lr': [11,0],
                    'Nx': [12,0],
                    'Ny': [12,0],
                    'R': [8,0],
                    'W0': [6,0],
                    'W90': [6,0] }
    
    [ETABS_MODEL.load_patterns.add(name,LoadPatternsVV[name][i], 
                                   self_weight_multiplier = 0, 
                                   add_analysis_case = False) for  i, name 
                                   in enumerate(LoadPatternsVV.keys())] 


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
    etabs_model = pytabs.EtabsModel(attach_to_instance=model_is_open, model_path=model_fp, backup=True)
    
    # set the model units to 
    etabs_model.set_present_units(etabs_model.eUnits.kgf_m_C)
    print(f"Units set to kgf_m_C")
    
    #Unlock model
    etabs_model.set_model_is_locked(False)

    #get existing loadcases and delete
    ExistLCs = etabs_model.load_cases.get_name_list()  
    if ExistLCs[0]!='':
        [etabs_model.load_cases.delete(name)   for name in ExistLCs if name != 'Modal' ]   
        [etabs_model.load_cases.delete('Modal')   for name in ExistLCs if name == 'Modal' ] 
        print(f"A total of {len(ExistLCs)} load cases were deleted: {ExistLCs}")


    #get existing load patterns and delete
    ExistLPs = etabs_model.load_patterns.get_name_list()  
    if ExistLPs[0]!='':
        [etabs_model.load_patterns.delete(name) for name in ExistLPs]    
        print(f"A total of {len(ExistLPs)} load patterns were deleted: {ExistLPs}")
        
    etabs_model.get_filename()
        
        
    #define load Patterns
    # assign_load_patterns_VV(etabs_model)
    
    print('done.')

    # exit ETABS if not attached
    if not model_is_open:
        etabs_model.exit_application()

    print(f"Number of target stories imported: ")

    print('\nWriting to excel...', end='')

    # write_batch_results(batch_wb, stories)
    print('done.')
    
    print('\nExtraction run complete, press any key to exit.')
    input('')


if __name__ == "__main__":
    xw.Book(WB_FN).set_mock_caller()
    main()