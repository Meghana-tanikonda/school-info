*** Settings ***
Documentation       Template robot main suite.

Library             RPA.Excel.Files
Library             RPA.Browser.Selenium    auto_close=${False}
Library             RPA.Tables
Library             RPA.FileSystem
#Library    RPA.Robocloud.Items
Library             RPA.Robocorp.WorkItems
Library             RPA.JSON
Library             Collections
Library             RPA.Outlook.Application

Suite Teardown      RPA.Outlook.Application.Quit Application
Task Setup          RPA.Outlook.Application.Open Application


*** Tasks ***
consumer School info
    TRY
        ${config}=    Load JSON from file
        ...    C:${/}Users${/}meghana.tanikonda${/}Documents${/}Robocorp${/}School info bots dna${/}config.json
        ${input_excel}=    Set Variable    ${config}[input]
        ${sheet_name}=    Set Variable    ${config}[sheet_name]
        ${Browser_url}=    Set Variable    ${config}[Browser]
        ${recipients}=    Set Variable    ${config}[recipients]
        ${Subject}=    Set Variable    ${config}[Subject]
        ${Body}=    Set Variable    ${config}[Body]
        ${code}=    For Each Input Work Item    load work items
        open Browser    ${Browser_url}    ${recipients}    ${Subject}    ${Body}
        ${Row}=    Set Variable    2
        FOR    ${input_data}    IN    @{code}
            TRY
                search the school code    ${input_data}
                ${handles}=    Get Window Handles
                get the all details    ${input_excel}    ${handles}    ${Row}    ${sheet_name}
                ${Row}=    Evaluate    ${Row}+1
            EXCEPT    message
                Log    Unable to getting the school details
            END
        END
    EXCEPT    message
        Log    Browser Not able to open
    END


*** Keywords ***
load work items
    ${work_items}=    Get Work Item Variables
    ${code}=    Set Variable    ${work_items}[code]
    RETURN    ${code}

open Browser
    [Arguments]    ${Browser_url}    ${recipients}    ${Subject}    ${Body}
    TRY
        Open Available Browser    ${Browser_url}
        Sleep    5s
        Maximize Browser Window
        Sleep    5s
    EXCEPT
        Send Exception mail    ${recipients}    ${Subject}    ${Body}
    END

search the school code
    [Arguments]    ${code}
    Input Text    SchoolCode    ${code}
    Click Button    //*[@id="SearchSchool"]
    ${handles}=    Get Window Handles
    Sleep    5s
    Switch Window    ${handles}[1]

 get the all details
    [Arguments]    ${input_excel}    ${handles}    ${Row}    ${sheet_name}
    ${school name}=    Get Text    css:body > center > h1
    ${School Address}=    Get Text    css:body > center > table > tbody > tr:nth-child(1) > td:nth-child(2)
    ${Phonenumber}=    Get Text    css:body > center > table > tbody > tr:nth-child(2) > td:nth-child(2)
    ${strength}=    Get Text    css:body > center > table > tbody > tr:nth-child(3) > td:nth-child(2)
    ${Prncipal Name }=    Get Text    css:body > center > table > tbody > tr:nth-child(4) > td:nth-child(2)
    ${no of teaching staff}=    Get Text    css:body > center > table > tbody > tr:nth-child(5) > td:nth-child(2)
    ${Number of Non-TeachingStaff}=    Get Text
    ...    css:body > center > table > tbody > tr:nth-child(6) > td:nth-child(2)
    ${Number of School buses}=    Get Text    css:body > center > table > tbody > tr:nth-child(7) > td:nth-child(2)
    ${School Playground}=    Get Text    css:body > center > table > tbody > tr:nth-child(8) > td:nth-child(2)
    ${Facilities}=    Get Text    css:body > center > table > tbody > tr:nth-child(9) > td:nth-child(2)
    ${School Accrediation}=    Get Text    css:body > center > table > tbody > tr:nth-child(10) > td:nth-child(2)
    ${School Hostel}=    Get Text    css:body > center > table > tbody > tr:nth-child(11) > td:nth-child(2)
    ${School Canteen}=    Get Text    css:body > center > table > tbody > tr:nth-child(12) > td:nth-child(2)
    ${School Stationary}=    Get Text    css:body > center > table > tbody > tr:nth-child(13) > td:nth-child(2)
    ${School Teaching method's}=    Get Text    css:body > center > table > tbody > tr:nth-child(14) > td:nth-child(2)
    ${School Timing}=    Get Text    css:body > center > table > tbody > tr:nth-child(15) > td:nth-child(2)
    ${School Achivements}=    Get Text    css:body > center > table > tbody > tr:nth-child(16) > td:nth-child(2)
    ${School Awards}=    Get Text    css:body > center > table > tbody > tr:nth-child(17) > td:nth-child(2)
    ${School Uniform}=    Get Text    css:body > center > table > tbody > tr:nth-child(18) > td:nth-child(2)
    ${School type}=    Get Text    css:body > center > table > tbody > tr:nth-child(19) > td:nth-child(2)
    Open Workbook    ${input_excel}
    Read Worksheet    ${sheet_name}
    Set Cell Value    ${Row}    B    ${school name}
    Set Cell Value    ${Row}    C    ${School Address}
    Set Cell Value    ${Row}    D    ${Phonenumber}
    Set Cell Value    ${Row}    E    ${strength}
    Set Cell Value    ${Row}    F    ${Prncipal Name }
    Set Cell Value    ${Row}    G    ${no of teaching staff}
    Set Cell Value    ${Row}    H    ${Number of Non-TeachingStaff}
    Set Cell Value    ${Row}    I    ${Number of School buses}
    Set Cell Value    ${Row}    J    ${School Playground}
    Set Cell Value    ${Row}    K    ${Facilities}
    Set Cell Value    ${Row}    L    ${School Accrediation}
    Set Cell Value    ${Row}    M    ${School Hostel}
    Set Cell Value    ${Row}    N    ${School Canteen}
    Set Cell Value    ${Row}    O    ${School Stationary}
    Set Cell Value    ${Row}    P    ${School Teaching method's}
    Set Cell Value    ${Row}    Q    ${School Timing}
    Set Cell Value    ${Row}    R    ${School Achivements}
    Set Cell Value    ${Row}    S    ${School Awards}
    Set Cell Value    ${Row}    T    ${School Uniform}
    Set Cell Value    ${Row}    U    ${School type}
    Save Workbook
    Close Window
    Sleep    5s
    Switch Window    ${handles}[0]
    Clear Element Text    //*[@id="SchoolCode"]

Send Exception mail
    [Arguments]    ${recipients}    ${Subject}    ${Body}
    Send Message    recipients=${recipients}    subject=${Subject}    body=${Body}
