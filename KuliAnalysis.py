""" KULI Analysis simulation interface """
import win32com.client
import re
import colorama

class KuliEvents(object):
    """ KuliEvents """

    def OnCheckForCancel(self):
        '''
        Event if fired if the user may interrupt the analysis. To interrupt, call
        Cancel() within this event.
        '''
        pass

    def OnEndOfOperatingPoint(self, operating_point):
        '''
        Event if fired at the end of each operating point.
        :param operating_point: the operting point number.
        '''
        pass

    def OnEndOfTimeStep(self, time_step, time):
        '''
        Event if fired at the end of each operating point.
        :param time_step: the number of the finished time step.
        :param time: actual value of the time.
        '''
        pass

    def OnError(self, function_name, message, additional_info, message_type):
        '''
        Event is fired if some error occured.
        :param function_name: name of the offending function.
        :param message: error message.
        :param additional_info: additional message information.
        :param message_type: error severity value.
        '''
        print(colorama.Back.RED + 'ERROR: ' + message + '\t' + additional_info + colorama.Style.RESET_ALL)

    def OnMessage(self, function_name, message, additional_info, message_type):
        '''
        Event forwards the messages from KULI.
        :param function_name: name of the function where the message is raised.
        :param message: message text.
        :param additional_info: additional message information.
        :param message_type: does currently not give any value.
        '''
        print(colorama.Fore.GREEN + message + '\t' + additional_info + colorama.Style.RESET_ALL)

    def OnNextIteration(self, iteration_number):
        '''
        Event if fired when the next iteration has begun.
        :param iteration_number: the number of the iteration.
        '''
        pass

    def OnNextTime(self, time_step_number, time):
        '''
        Event if fired when the next time step has been reached. Requires the
        simulation to be started either by using RunAnalysis() or SimualteOperatingPoint(...).
        :param time_step_number: the number of the time step.
        :param time_step_number: actual value of the time.
        '''
        pass


class KuliAnalysis(object):
    """ Python wrapper for KULIAnalysis2Ctr"""

    def __init__(self, versionstring, kuli_event_handler=None):
        """ Constructur
        :param versionstring: e.g. 11.1
        """
        colorama.init()
        server_name = 'KuliAnalysis2.KuliAnalysisCtr2.'
        versionnumber = int(re.search('^[0-9]+', versionstring).group())
        if versionnumber is not None and versionnumber >= 13:
            server_name = 'KuliAnalysisServer.KuliAnalysisCtr2.'

        versionid = server_name + versionstring

        event_handler = kuli_event_handler
        if event_handler is None:
            event_handler = KuliEvents
        self.__kuli = win32com.client.DispatchWithEvents(versionid, event_handler)


    # ACAdjustmentMode property
    def __get_ac_adjustment_mode(self):
        return self.__kuli.ACAdjustmentMode()

    def __set_ac_adjustement_mode(self, val):
        self.__kuli.ACAdjustmentMode = val

    ac_adjustment_mode = property(__get_ac_adjustment_mode, __set_ac_adjustement_mode)

    # AddToBatchList
    def add_to_batch_list(self, file_name):
        ''' add_to_batch_list '''
        self.__kuli.AddToBatchList(file_name)

    # AnalysisLogEvents property
    def __get_analysis_log_events(self):
        return self.__kuli.AnalysisLogEvents

    def __set_analysis_log_events(self, val):
        self.__kuli.AnalysisLogEvents = val

    analysis_log_events = property(__get_analysis_log_events, __set_analysis_log_events)

    # AnalysisLogFile property
    def __get_analysis_log_file(self):
        return self.__kuli.AnalysisLogFile

    def __set_analysis_log_file(self, val):
        self.__kuli.AnalysisLogFile = val

    analysis_log_file = property(__get_analysis_log_file, __set_analysis_log_file)

    # BatchMode property
    def __get_batch_mode(self):
        return self.__kuli.BatchMode

    def __set_batch_mode(self, val):
        self.__kuli.BatchMode = val

    batchMode = property(__get_batch_mode, __set_batch_mode)

    # Cancel
    def cancel(self):
        """ cancel """
        return self.__kuli.Cancel()

    # CleanUp
    def clean_up(self):
        """ clean_up """
        return self.__kuli.CleanUp()

    def get_com_codes(self, is_input_com_object):
        '''
        returns either all input or output COM object
        :param is_input_com_object: True for input COM objects, otherwise False
        '''
        return self.__kuli.GetCOMCodes(is_input_com_object)

    def get_com_value_by_id(self, com_id):
        '''
        returns the numerical value of a COM with the specified ID
        :param com_id: ID of the COM object
        '''
        return self.__kuli.GetCOMValueByID(com_id)

    def get_com_value_by_id_as_string(self, com_id):
        '''
        returns the a string value of a COM with the specified ID
        :param com_id: ID of the COM object
        '''
        return self.__kuli.GetCOMValueByIDAsString(com_id)

    # GetCOMValueByIDAsString2
    def get_com_value_by_id_as_string2(self, com_id):
        '''
        returns the a string value of a COM with the specified ID
        :param com_id: ID of the COM object
        '''
        return self.__kuli.GetCOMValueByIDAsString2(com_id)

    #GetInputCOMUnitByID
    def get_input_com_unit_by_id(self, com_id):
        '''
        returns the unit of a input COM with the specified ID
        :param com_id: ID of the COM object
        '''
        return self.__kuli.GetInputCOMUnitByID(com_id)


    #GetOutputCOMUnitByID
    def get_output_com_unit_by_id(self, com_id):
        '''
        returns the unit of a output COM with the specified ID
        :param com_id: ID of the COM object
        '''
        return self.__kuli.GetOutputCOMUnitByID(com_id)    

    # GetComponentDescription
    def get_component_description(self, component_type_id):
        '''
        returns the full name of a component with the specified ID
        :param component_type_id: ID of the component type
        '''
        return self.__kuli.GetComponentDescription(component_type_id)

    # GetComponentConnectorUnit
    def get_component_connector_unit(self, component_type_id, connector_id):
        '''
        returns the unit for a specified component connector
        :param component_type_id: ID of the component type
        :param connector_id: ID of the connector
        '''
        return self.__kuli.GetComponentConnectorUnit(component_type_id, connector_id)

    # GetConnectorUnit
    def get_connector_unit(self, connector_id):
        '''
        returns the unit for a specified connector
        :param connector_id: ID of the connector
        '''
        return self.__kuli.GetConnectorUnit(connector_id)

    # GetErrorInfo
    ######## This does not work ###########
    # strings are immuteable
    # def get_error_info(self):
    #     '''
    #     returns the number of errors and errors texts
    #     '''
    #     number_of_errors = self.__kuli.GetConnectorUnit(errors)
    #     return number_of_errors, errors

    # GetValue
    def get_value(self, component_id, connector_id):
        """
        returns the double value of a sensor or actuator
        :param component_id: id of the component
        :param connector_id: id of the sensor or actuator
        """
        return self.__kuli.GetValue(component_id, connector_id)

    # GetValueAsStr
    def get_value_as_str(self, component_id, connector_id):
        """
        returns the string value of a sensor or actuator
        :param component_id: ID of the component
        :param connector_id: ID of the sensor or actuator
        """
        return self.__kuli.GetValueAsStr(component_id, connector_id)

    # GetValueUnit
    def get_value_unit(self, component_id, connector_id, unit):
        """
        returns the double value of a sensor or actuator converted to the specified unit
        :param component_id: ID of the component
        :param connector_id: ID of the sensor or actuator
        :param unit: unit of the value
        """
        return self.__kuli.GetValueUnit(component_id, connector_id, unit)

    # GetValueUnitAsStr
    def get_value_unit_as_str(self, component_id, connector_id, unit):
        """
        returns the string value of a sensor or actuator converted to the specified unit
        :param component_id: ID of the component
        :param connector_id: ID of the sensor or actuator
        :param unit: unit of the value
        """
        return self.__kuli.GetValueUnitAsStr(component_id, connector_id, unit)

    # GetWarningInfo
    ######## This does not work ###########
    # strings are immuteable
    # def get_warning_info(self):
    #     '''
    #     returns the number of warnings the warning texts
    #     '''
    #     number_of_warnings = self.__kuli.GetWarningInfo(warning)
    #     return number_of_warnings, warning

    # GetVersionID
    def get_version_id(self):
        '''
        returns the version id of KULI
        '''
        return self.__kuli.GetVersionID()

    # GetVersionStr
    def get_version_str(self):
        '''
        returns the version string of KULI
        '''
        return self.__kuli.GetVersionStr()

    # Initialize
    def initialize(self):
        '''
        initalizes the specified scs file
        '''
        if self.__kuli is None:
            return False
        return self.__kuli.Initialize()

    # IsFinished
    def is_finished(self):
        '''
        Returns True if the computation of the current simultion is finished,
        otherwise False.
        '''
        return self.__kuli.IsFinished()

    # IsNextTimeStep
    def is_next_time_step(self):
        '''
        Returns True is the compution of the current time step succeeded and
        an additional time step is left, otherwise false.
        '''
        return self.__kuli.IsNextTimeStep()

    # ListComponents
    def list_components(self, filter):
        '''
        Returns a string containing a list of components of the KULI model, separated
        by spaces.
        :param filter: string to filter data
        '''
        return self.__kuli.ListComponents(filter)

    # ListComponentTypes
    def list_component_types(self):
        '''
        Returns a string containg the IDs of all KULI component types seperated with blanks.
        '''
        return self.__kuli.ListComponentTypes()

    # ListConnectors
    def list_connectors(self, component_type, connector_type):
        '''
        Returns a string with the IDs of all connectors of a component type separated
        by spaces.
        :param connector_type: can be "SA" to return all connectors,
        "S" for sensors and "A" for actuators
        '''
        return self.__kuli.ListConnectors(component_type, connector_type)

    # IsNextTimeStep
    def next_kuli_iteration(self):
        '''
        Starts the computation of the next iteration of the current operating point.
        Returns True is the calculation was successful.
        '''
        return self.__kuli.NextKULIIteration()

    # ObjectExists
    def object_exists(self, com_object_name, in_or_out):
        '''
        Returns true if the object could be found.
        :param com_object_name: the name of the COM object
        :param in_or_out: can be "In" (Input COM object) or "Out" (Output COM object).
        '''
        return self.__kuli.ObjectExists(com_object_name, in_or_out)

    # ResetCurrentSimulationStep
    def reset_current_simulation_step(self):
        '''
        Sets the IsFinished flag to false and forces a recalculation of the current
        simulation step.
        Returns true if the simulation was successful.
        '''
        return self.__kuli.ResetCurrentSimulationStep()

    # RunAnalysis
    def run_analysis(self):
        '''
        Performs a complete run of a cooling system simulation.
        Returns True if the simulation was successful.
        '''
        return self.__kuli.RunAnalysis()

    # RunOptimization
    def run_optimization(self):
        '''
        Performs an optimization.
        Returns True if the simulation was successful.
        '''
        return self.__kuli.RunOptimization()

    # RunParameterVariation
    def run_parameter_variation(self):
        '''
        Performs a parameter variation simulation.
        Returns True if the simulation was successful.
        '''
        return self.__kuli.RunParameterVariation()

    # RunMonteCarlo
    def run_monte_carlo(self, no_of_samples):
        '''
        Performs a monte carlo simulation.
        Returns True if the simulation was successful.
        :param no_of_samples: number of samples.
        '''
        return self.__kuli.RunMonteCarlo(no_of_samples)

    # KuliFileName property
    def __get_kuli_file_name(self):
        return self.__kuli.KuliFileName

    def __set_kuli_file_name(self, val):
        self.__kuli.KuliFileName = val

    kuli_file_name = property(__get_kuli_file_name, __set_kuli_file_name)

    # ResultFileName
    def __get_result_file_name(self):
        return self.__kuli.ResultFileName

    def __set_result_file_name(self, val):
        self.__kuli.ResultFileName = val

    result_file_name = property(__get_result_file_name, __set_result_file_name)

    # SetCOMValueByID
    def set_com_value_by_id(self, component_id, value):
        '''
        Sets the double value of an input COM object.
        Returns True if the value was set successfully.
        :param component_id: ID of the COM object.
        :param value: value to be set.
        '''
        return self.__kuli.SetCOMValueByID(component_id, value)

    # SetCOMValueByIDAsStr
    def set_com_value_by_id_as_str(self, component_id, value):
        '''
        Sets the string value of an input COM object.
        Returns True if the value was set successfully.
        :param component_id: ID of the COM object.
        :param value: value to be set.
        '''
        return self.__kuli.SetCOMValueByIDAsStr(component_id, value)

    # SetCOMValueByIDAsStr2
    def set_com_value_by_id_as_str2(self, component_id, value):
        '''
        Sets the string value of an input COM object.
        Returns True if the value was set successfully.
        :param component_id: ID of the COM object.
        :param value: value to be set.
        '''
        return self.__kuli.SetCOMValueByIDAsStr2(component_id, value)

    # SetValue
    def set_value(self, component_id, actuator_id, value):
        '''
        Sets the double value of an actuator for the specified component.
        :param component_id: ID of the component.
        :param connector_id: ID of the actuator.
        :param value: value to be set.
        '''
        self.__kuli.SetValue(component_id, actuator_id, value)

    # SetValueAsStr
    def set_value_as_str(self, component_id, actuator_id, value):
        '''
        Sets the string value of an actuator for the specified component.
        :param component_id: ID of the component.
        :param connector_id: ID of the actuator.
        :param value: value to be set.
        '''
        self.__kuli.SetValueAsStr(component_id, actuator_id, value)

    # SetValueUnit
    def set_value_unit(self, component_id, actuator_id, unit, value):
        '''
        Sets the double value of an actuator for the specified component.
        :param component_id: ID of the component.
        :param connector_id: ID of the actuator.
        :param unit: unit of the value to be set.
        :param value: value to be set.
        '''
        self.__kuli.SetValueUnit(component_id, actuator_id, unit, value)

    # SetValueUnitAsStr
    def set_value_unit_as_str(self, component_id, actuator_id, unit, value):
        '''
        Sets the string value of an actuator for the specified component.
        :param component_id: ID of the component.
        :param connector_id: ID of the actuator.
        :param unit: unit of the value to be set.
        :param value: value to be set.
        '''
        self.__kuli.SetValueUnitAsStr(component_id, actuator_id, unit, value)

    # SimulateOperatingPoint
    def simulate_operating_point(self, operating_point):
        '''
        Simululates the specified operating point.
        :param operating_point: operating point to be simulated. A value
        of 0 simulates all operating points.
        '''
        return self.__kuli.SimulateOperatingPoint(operating_point)

    # StartAnalysis
    def start_analysis(self):
        '''
        Initializes the system and starts the analysis.
        '''
        return self.__kuli.StartAnalysis()

    # WriteResults
    def __get_write_results(self):
        return self.__kuli.WriteResults

    def __set_write_results(self, val):
        self.__kuli.WriteResults = val

    write_results = property(__get_write_results, __set_write_results)
