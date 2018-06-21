from KuliAnalysis import KuliAnalysis
import colorama

# starts and initializes kuli
def start_kuli(kuli_file_name):
    KULI = KuliAnalysis('13')
    KULI.kuli_file_name = kuli_file_name
    KULI.write_results = False

    if KULI.initialize() == 0:
        print(colorama.Back.RED 
            + "ERROR: KULI could not be initalized!"
            + colorama.Style.RESET_ALL)
        return None
    
    return KULI

# clean up and destroy KULI
def cleanup_kuli(kuli):
    kuli.clean_up()
    del kuli
#############################################################################

#############################################################################
# component connector helper functions

# extracts the KULI component id from a string
# sample: returns "1.Rad" from "1.Rad.ExitTempIM"
def extract_component_id(component_connector):
    number_of_delim = component_connector.count(".")
    
    if number_of_delim != 1 and number_of_delim != 2:
        return ""
    
    return component_connector.rsplit('.', 1)[0]

# extracts the KULI component type from a string
# sample: returns "Rad" from "1.Rad.ExitTempIM"
def extract_component_type(component_connector):
    number_of_delim = component_connector.count(".")
    
    if number_of_delim == 1:
        return component_connector.split('.')[0]
    elif number_of_delim == 2:
        return component_connector.split(".")[1] # take "Rad" from "1.Rad.ExitTempIM[K]"
    else:
        return ""

# extracts the KULI connector id from a string
# sample: returns "ExitTempIM" from "1.Rad.ExitTempIM"
def extract_connector_id(component_connector, keep_unit:bool=True):
    connector_id = component_connector

    if not keep_unit:
        connector_id = component_connector.split("[")[0]  # remove unit definition from string

    number_of_delim = connector_id.count(".")
    
    if number_of_delim != 1 and number_of_delim != 2:
        return ""

    return connector_id.rsplit('.', 1)[1] # take right most enty from string

# validates a list of component connectors
def validate_connectors(component_connectors:list, connectors_are_inputs:bool, kuli_file_name:str):
    KULI = start_kuli(kuli_file_name) # type: KuliAnalysis
    if KULI is None:
        return False

    component_ids = KULI.list_components("*").split()
    if len(component_ids) == 0:
        print(colorama.Back.RED 
            + "ERROR: Error getting components list from KULI."
            + colorama.Style.RESET_ALL)
        return False

    valid = True

    for component_connector in component_connectors: 
        component_id = extract_component_id(component_connector)

        if component_id in component_ids:
            component_type = extract_component_type(component_connector)
            if component_type == "":
                print(colorama.Back.RED 
                    + "ERROR: Invalid component type: %s" % component_connector
                    + colorama.Style.RESET_ALL)
                valid = False
            else:
                connectors = KULI.list_connectors(component_type , "A" if connectors_are_inputs else "S").split()
                connector = extract_connector_id(component_connector, False)
                if connector in connectors:
                    print(colorama.Fore.GREEN 
                        + "Info:\tConnector successfully validated: \'%s\'" % component_connector
                        + colorama.Style.RESET_ALL)
                else:
                    print(colorama.Back.RED 
                        + "ERROR: Invalid connector ID: %s" % component_connector
                        + colorama.Style.RESET_ALL)
                    valid = False

        else:
            if component_id == "COM": # COM objects cannot be validated
                print(colorama.Fore.GREEN 
                    + "INFO:\tCOM objects cannot be validated: %s" % component_connector
                    + colorama.Style.RESET_ALL)
            else:
                print(colorama.Back.RED 
                    + "ERROR: Invalid component ID: %s" % component_connector
                    + colorama.Style.RESET_ALL)
                valid = False

    cleanup_kuli(KULI)

    return valid
#############################################################################
