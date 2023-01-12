*** Settings ***
Documentation       Template robot main suite.

Library             RPA.Excel.Files
Library             RPA.Tables
Library             RPA.Robocorp.WorkItems
Library             Collections
Library             RPA.JSON
Library             RPA.Outlook.Application
Library             RPA.FileSystem

Suite Teardown      RPA.Outlook.Application.Quit Application
Task Setup          RPA.Outlook.Application.Open Application


*** Tasks ***
Producer school info
    TRY
        ${config}=    Load JSON from file
        ...    C:${/}Users${/}meghana.tanikonda${/}Documents${/}Robocorp${/}School info bots dna${/}config.json
        ${input_excel}=    Set Variable    ${config}[input]
        ${sheet_name}=    Set Variable    ${config}[sheet_name]
        ${recipients}=    Set Variable    ${config}[recipients]
        ${Subject}=    Set Variable    ${config}[Subject]
        ${Body}=    Set Variable    ${config}[Body]
        ${table}=    Read excel    ${input_excel}    ${sheet_name}    ${recipients}    ${Subject}    ${Body}
        uploading WorkItems    ${table}
    EXCEPT    message
        Log    Excel Not found
    END


*** Keywords ***
Read excel
    [Arguments]    ${input_excel}    ${sheet_name}    ${recipients}    ${Subject}    ${Body}
    ${file_exist}=    Does File Exist    ${input_excel}
    IF    ${file_exist} == ${True}
        Open Workbook    ${input_excel}
        ${table}=    Read Worksheet As Table    ${sheet_name}    ${True}
        RETURN    ${table}
    ELSE
        Send Exception mail    ${recipients}    ${Subject}    ${Body}
    END

uploading WorkItems
    [Arguments]    ${table}
    FOR    ${row}    IN    @{table}
        ${school_code}=    Set Variable    ${row}[School Code]
        ${dict}=    Create Dictionary    code=${school_code}
        Create Output Work Item    variables=${dict}    save=True
    END

Send Exception mail
    [Arguments]    ${recipients}    ${Subject}    ${Body}
    Send Message    recipients=${recipients}    subject=${Subject}    body=${Body}
