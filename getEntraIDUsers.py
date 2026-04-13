import msal
import json
import requests

# import master
import pandas as pd
import os
from dotenv import load_dotenv

############################################################################################################
# get data from Entra ID (Active Directory) using Microsoft Graph API
############################################################################################################


def get_graph_access_token():
    load_dotenv()
    tenantID = "0c16d39e-94a9-4415-8049-e2922dbc1a2e"  # ambaflex
    authority = "https://login.microsoftonline.com/" + tenantID
    clientID = os.getenv("graph_client_id")
    clientSecret = os.getenv("graph_client_secret")
    scope = ["https://graph.microsoft.com/.default"]
    app = msal.ConfidentialClientApplication(
        clientID, authority=authority, client_credential=clientSecret
    )
    access_token = app.acquire_token_for_client(scopes=scope)
    return access_token["access_token"]


def get_all_users_of_group(token, group_name):
    # get group id of the group that starts with GS_NL_CTX_ShareFile
    url = (
        "https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName, '"
        + group_name
        + "')"
    )
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    response = requests.get(url, headers=headers)
    data = json.loads(response.text)
    sharefileADgroupID = data["value"][0]["id"]

    # get all members of the group
    url = "https://graph.microsoft.com/v1.0/groups/" + sharefileADgroupID + "/members"
    response = requests.get(url, headers=headers)
    data = json.loads(response.text)
    all_users = data["value"]
    while "@odata.nextLink" in data:
        response = requests.get(data["@odata.nextLink"], headers=headers)
        data = response.json()
        all_users.extend(data["value"])

    # for member in all_users:
    #    print(member["id"
    # ],member["userPrincipalName"])
    return all_users


def departmentfound(splitdepartmentfilter, userdepartment):
    found = False
    for departmentfilter in splitdepartmentfilter.split(","):
        if departmentfilter in userdepartment:
            found = True
    return found


def get_all_Entra_Users(token, departmentfilterlist, lost_and_found):
    url = "https://graph.microsoft.com/v1.0/users?$count=true&$select=id,displayName,userPrincipalName,department,jobTitle,accountEnabled"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    response = requests.get(url, headers=headers)
    data = response.json()
    all_users = data["value"]

    while "@odata.nextLink" in data:
        response = requests.get(data["@odata.nextLink"], headers=headers)
        data = response.json()
        all_users.extend(data["value"])

    filtered_users = [
        user
        for user in all_users
        if departmentfound(departmentfilterlist, str(user["department"]))
        or user["userPrincipalName"].lower() in lost_and_found
        or True
    ]
    return all_users, filtered_users


############################################################################################################
# get data from citrix sharefile
############################################################################################################


def loginCitrixShareFile():
    load_dotenv()
    loginURL = "https://AmbaflexManufacturingBV.sharefile.com/oauth/token"
    login_data = {
        "grant_type": "password",
        "username": os.getenv("user_name"),
        "password": os.getenv("password"),
        "client_id": os.getenv("client_id"),
        "client_secret": os.getenv("client_secret"),
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    r = requests.post(loginURL, headers=headers, data=login_data)
    if r.status_code != 200:
        print("Login failed")
        print("Status code: ", r.status_code)
        print("Response: ", r.text)
        return None

    jsonResponse = json.loads(r.text.encode("utf8"))
    token = jsonResponse["access_token"]
    header = {
        "Accept": "application/json",
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json",
    }

    return header


def get_sharefile_user_id_by_email(email, header_with_token):
    url = f"https://AmbaflexManufacturingBV.sf-api.eu/sf/v3/Users?emailAddress={email}"
    response = requests.get(url, headers=header_with_token)

    if response.status_code == 200:
        data = response.json()
        user_id = data.get("Id", "")
        return user_id
    else:
        output = f"Failed to retrieve data 2: {response.status_code}"
        return "None"


def get_full_list_of_employees(header_with_token):
    url = f"https://AmbaflexManufacturingBV.sf-api.eu/sf/v3/Accounts/Employees"
    response = requests.get(url, headers=header_with_token)

    data = json.loads(response.text)
    employees = data["value"]

    count = data.get("odata.count", "")

    while "@odata.nextLink" in data:
        response = requests.get(data["@odata.nextLink"], headers=header_with_token)
        data = response.json()
        employees.extend(data["value"])

    return employees, count


# here are de details of all the sharefile of the employees of the department that have a sharefile account
def get_home_folder_by_user_id(user_id, header_with_token):

    if user_id == None:
        return None

    url = f"https://AmbaflexManufacturingBV.sf-api.eu/sf/v3/Users({user_id})/HomeFolder"
    response = requests.get(url, headers=header_with_token)
    if response.status_code == 200:
        data = response.json()
        # filecount = data.get("FileCount", "")
        FileSizeBytes = data.get("FileSizeBytes", "")
        FileSizeMBytes = round(FileSizeBytes / 1024 / 1024, 4)
        # print(data)
        return FileSizeMBytes
    else:
        print(url)
        print(f"Failed to retrieve data 3: {response.status_code}")
        return -1


def get_user_id_by_email(email, header_with_token):
    url = f"https://AmbaflexManufacturingBV.sf-api.eu/sf/v3/Users?emailAddress={email}"
    response = requests.get(url, headers=header_with_token)

    if response.status_code == 200:
        data = response.json()
        user_id = data.get("Id", "")
        return user_id, None
    else:
        output = f"Failed to retrieve data 4: {response.status_code}"
        return None, output


def get_all_shared_folders_by_user_id(user_id, header_with_token):
    url = f"https://AmbaflexManufacturingBV.sf-api.eu/sf/v3/Users({user_id})/AllSharedFolders"

    response = requests.get(url, headers=header_with_token)
    if response.status_code == 200:
        data = response.json()
        listSharedFolder = data.get("value", "")
        return listSharedFolder
    else:
        print(f"Failed to retrieve data 1: {response.status_code}")
        return None


def extract_all_has_vroot(data):
    folders = []
    for item in data:
        # print(item)
        # info = item.get('Info', {})
        CreatorShortName = item.get("CreatorNameShort", None)
        name = item.get("Name", None)
        folders.append([CreatorShortName, name])
        folders.sort()
    return folders


def print_shared_folders_of_one_employee(email, header_with_token):
    output = "-" * 80
    user_id, useroutput = get_user_id_by_email(email, header_with_token)
    output += f"\nUser : {email}, {user_id}\n"

    if user_id:
        fileSize = get_home_folder_by_user_id(user_id, header_with_token)
        output += f"Home Folder Size: {fileSize} MB\n"
        shared_folders = get_all_shared_folders_by_user_id(user_id, header_with_token)
        if shared_folders:
            # print(f"Shared Folders: {shared_folders}")
            sharedFoldersNames = extract_all_has_vroot(shared_folders)
            output += "{:<20}; {:<60}".format("Creator", "Folder Name") + "\n"
            output += "{:<20}; {:<60}".format("-------", "-----------") + "\n"

            for folder in sharedFoldersNames:
                output += "{:<20}; {:<60}".format(folder[0], folder[1]) + "\n"
        else:
            output += "No shared folders found"

    else:
        output += useroutput + "\n"
        output += "No user found"
    output += "\n"
    return output


def departmentCitrixShareFileUsages(employees, header_with_token, department_filter):
    f = open(
        "DepartmentCitrixShareFileUsages_" + department_filter + ".txt",
        "w",
        encoding="utf-8",
    )
    for employee in employees:
        output = print_shared_folders_of_one_employee(employee, header_with_token)
        f.write(output + "\n")
        print(output)
        f.write("\n")
        print("")
    f.close()


############################################################################################################
# export to excel the results
############################################################################################################


def create_excel_with_users(filtered_users, department_filter):
    # Prepare data for the first tab
    users_data = []
    NoYes = ["No", "Yes"]
    YesNo = ["Yes", "No"]
    line = 1
    for user in filtered_users:
        line = line + 1
        users_data.append(
            [
                NoYes[user["accountEnabled"]],
                YesNo[user["IsSharefileEnabled"]],  # flipped!
                NoYes[user["InShareFileGroup"]],
                "None" if user["sharefile_id"] is None else "Yes",
                user["userPrincipalName"],
                user["Name"],
                str(user["jobTitle"]),
                str(user["department"]),
                ("" if user["HomeFolderSizeMB"] is None else user["HomeFolderSizeMB"]),
                f"=COUNTIF(DepartmentUsages!A:A,Users!E{line})",
            ]
        )

    # Create a DataFrame for the first tab
    df_users = pd.DataFrame(
        users_data,
        columns=[
            "AccountEnabled",
            "InShareFileEnabled",
            "InShareFileGroup",
            "SharefileID",
            "UserPrincipalName",
            "Name",
            "JobTitle",
            "Department",
            "HomeFolderSizeMB",
            "SharedFoldersCount",
        ],
    )

    # Create an Excel writer object and write the DataFrame to the first tab
    departmentfilename = department_filter.replace(",", "_")[:30]
    with pd.ExcelWriter(
        f"DepartmentCitrixShareFileUsages_{departmentfilename}.xlsx"
    ) as writer:
        df_users.to_excel(writer, sheet_name="Users", index=False)

    print(
        f"Excel file 'DepartmentCitrixShareFileUsages_{department_filter}.xlsx' created successfully."
    )


def add_department_usages_to_excel(filtered_users, sharefile_header, department_filter):
    # Prepare data for the second tab
    department_usages_data = []
    line = 1
    for user in filtered_users:
        if user["sharefile_id"] is not None:
            email = user["userPrincipalName"]
            user_id, _ = get_user_id_by_email(email, sharefile_header)
            if user_id:
                fileSize = user[
                    "HomeFolderSizeMB"
                ]  # get_home_folder_by_user_id(user_id, sharefile_header)
                shared_folders = get_all_shared_folders_by_user_id(
                    user_id, sharefile_header
                )
                if shared_folders:
                    sharedFoldersNames = extract_all_has_vroot(shared_folders)
                    for folder in sharedFoldersNames:
                        line = line + 1
                        department_usages_data.append(
                            [
                                email,
                                fileSize,
                                folder[0],
                                folder[1],
                                f"=VLOOKUP(A{line},Users!E:H,4,FALSE)",
                                f"=VLOOKUP(A{line},Users!E:K,6,FALSE)",
                            ]
                        )
                else:  # no shared folders found
                    line = line + 1
                    department_usages_data.append([email, fileSize, "", ""])

    # Create a DataFrame for the second tab
    df_department_usages = pd.DataFrame(
        department_usages_data,
        columns=[
            "Email",
            "HomeFolderSizeMB",
            "Creator",
            "FolderName",
            "Department",
            "SharedFoldersCount",
        ],
    )

    # Append the DataFrame to the existing Excel file
    departmentfilename = department_filter.replace(",", "_")[:30]
    with pd.ExcelWriter(
        f"DepartmentCitrixShareFileUsages_{departmentfilename}.xlsx",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        df_department_usages.to_excel(
            writer, sheet_name="DepartmentUsages", index=False
        )

    print(
        f"Excel file 'DepartmentCitrixShareFileUsages_{department_filter}.xlsx' updated with department usages."
    )


def add_totals_to_excel(
    filtered_users, sharefile_users, department_filter, sharefile_employees
):
    # Calculate totals
    total_users = len(filtered_users)
    total_sharefile_employees = len(sharefile_employees)
    total_sharefile_group_users = len(sharefile_users)
    total_in_sharefile_group = sum(user["InShareFileGroup"] for user in filtered_users)
    total_not_account_enabled_in_sharefile_group = sum(
        not user["accountEnabled"] and user["InShareFileGroup"]
        for user in filtered_users
    )
    total_in_citrix_sharefile_account = sum(
        1 for user in filtered_users if user["sharefile_id"] is not None
    )

    # Prepare data for the third tab
    totals_data = [
        ["Total users in Entra ID with department", department_filter, total_users],
        [
            "Total accounts in Sharefile (license)",
            "",
            total_sharefile_employees,
        ],
        [
            "Total users in Sharefile group",
            "GS_NL_CTX_ShareFile",
            total_sharefile_group_users,
        ],
        [
            "Total users in Entra ID with department and in Sharefile group",
            department_filter,
            total_in_sharefile_group,
        ],
        [
            "Total users in Entra ID with department and in Sharefile group and account disabled",
            department_filter,
            total_not_account_enabled_in_sharefile_group,
        ],
        [
            "Total users in Entra ID with department and has a Sharefile license",
            department_filter,
            total_in_citrix_sharefile_account,
        ],
    ]

    # Create a DataFrame for the third tab
    df_totals = pd.DataFrame(totals_data, columns=["Description", "Filter", "Count"])
    departmentfilename = department_filter.replace(",", "_")[:30]
    # Append the DataFrame to the existing Excel file
    with pd.ExcelWriter(
        f"DepartmentCitrixShareFileUsages_{departmentfilename}.xlsx",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        df_totals.to_excel(writer, sheet_name="Totals", index=False)

    print(
        f"Excel file 'DepartmentCitrixShareFileUsages_{department_filter}.xlsx' updated with totals."
    )


def add_open_departments_to_excel(departments):

    # Create a DataFrame for the fourth tab
    department_todo_data = []
    for department in departments:
        department_todo_data.append(
            [department["description"], department["user_count"], department["name"]]
        )

    df_totals = pd.DataFrame(
        department_todo_data, columns=["Department", "UserCount", "Users"]
    )
    departmentfilename = department_filter.replace(",", "_")[:30]
    # Append the DataFrame to the existing Excel file
    with pd.ExcelWriter(
        f"DepartmentCitrixShareFileUsages_{departmentfilename}.xlsx",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        df_totals.to_excel(writer, sheet_name="Other Departments", index=False)

    print(
        f"Excel file 'DepartmentCitrixShareFileUsages_{department_filter}.xlsx' updated with totals."
    )


############################################################################################################
# main
############################################################################################################


if __name__ == "__main__":
    # lost_and_found = ["mhoogeveen@ambaflex.com", "tklijzing@ambalfex.com"]
    # department_filter = "HR"

    department_filter = "Fin"
    lost_and_found = ["ax@ambaflex.com"]

    department_filter = "IT"

    department_filter = "Spa"
    lost_and_found = ["spare@ambaflex.com"]

    department_filter = "HR"
    lost_and_found = ["mhoogeveen@ambaflex.com", "tklijzing@ambalfex.com"]

    department_filter = "One_user_departments"

    department_filter = "engin"

    department_filter = "P_M_O"
    lost_and_found = [
        "KMeilink@Ambaflex.com",
        "JSmallegange@Ambaflex.com",
        "jvlasveld@ambaflex.com",
    ]

    lost_and_found = []

    department_filter = "SA_Salesdrawnings"
    lost_and_found = ["Salesdrawings@ambaflex.com"]

    department_filter = "Serv,Spa"
    lost_and_found = ["x"]

    department_filter = "Fac"

    department_filter = "Serv,Spa"
    lost_and_found = []
    department_filter = "Assembly, bal Order Man, Engineering, Warehouse, Construction"
    department_filter = "AmbaFlex Integration Service,Assembly,Coating RO,Construction,Global Application Engineering,Global Operations Support,Global Order Management,Hands-on Development,IBD,Industry Sub Chains,Injection Moulding,Innovation Development,Interim Inkoper,Logiflexx Conveyor Solutions,Materials Management,Medanva Employees,MES,OMD - PDM,Order Engineering CN Employees,Order Management,Product Engineering,Production,Project Management,Quality Engineering,Replenishing,Sector Business Administration,Sector Innovation,Sector Operations,Sector Sales,Support Desk,Tactical Purchasing US Employees,Warehouse"
    department_filter = "AmbaFlex Integration Service, Coating,Construction, Global Application Engineering, Global Operations Support, Global Order Management, Hands-on Development, IBD, Industry Sub Chains, Injection Moulding,Innovation Development,Interim Inkoper, Logiflexx Conveyor Solutions, Materials Management, Medanva Employees, MES, OMD, Order Engineering , Order Management, Product Engineering, Production, Project Management, Quality Engineering, Replenishing, Sector Business Administration, Sector Innovation, Sector Operations, Sector Sales, Support Desk, Tactical Purchasing, Warehouse"
    department_filter = "Integration, Coating,Construction, Global Application Engineering, Global Operations Support, Global Order Management, Hands, IBD, Industry Sub Chains, Injection Moulding,Innovation Development,Interim Inkoper, Logiflexx Conveyor Solutions, Materials Management, Medanva Employees, MES, OMD, Order Engineering , Order Management, Product Engineering, Production, Project Management, Quality Engineering, Replenishing, Sector Business Administration, Sector Innovation, Sector Operations, Sector Sales, Support Desk, Tactical Purchasing, Warehouse"
    department_filter = "Integration, Coating, Construction, Global Application Engineering, Global Operations Support, Global Order Management, Hands, IBD, Industry Sub Chains, Injection Moulding,Innovation Development,Interim Inkoper, Logiflexx Conveyor Solutions, Materials Management, Medanva Employees, MES, OMD, Order Engineering , Order Management, Product Engineering, Production, Project Management, Quality Engineering, Replenishing, Sector Business Administration, Sector Innovation, Sector Operations, Sector Sales, Support Desk, Tactical Purchasing, Warehouse"
    department_filter = "Integration, Coating, Construction, Engineering, Operations Support, Order Management, Hands, IBD, Industry Sub Chains, Injection Moulding,Innovation Development,Interim Inkoper"

    department_filter = "IBD"
    department_filter = (
        "Sales,Global Application Engineering,IBD,Logiflexx Conveyor Solutions"
    )
    department_filter = "SA_Salesdrawnings"
    department_filter = "Hands"
    department_filter = "Customer Service AMCAS,Field Service AMCAS Employees,Spare Parts Warehouse US Shift 1 Employees"

    department_filter = "Mar"
    department_filter = "TEAM_Hero,R&D,Innovation Development,Product Engineering"
    lost_and_found = ["AHagg@Ambaflex.com"]
    department_filter = "Coating,Construction,Operations S,Order Management,Hands,Injection Moulding,Interim Inkoper,Materials Management,MES,OMD,Order Engineering,Order Management,Production,Quality Engineering,Replenishing,Sector Operations,Tactical Purchasing,Warehouse,Assembly"
    department_filter = "Serv,Spa"
    department_filter = "MT"
    lost_and_found = [
        "sdejager@ambaflex.com",
        "Ammeraal@ambaflex.com",
        "JVlasveld@Ambaflex.com",
        "wbalk@ambaflex.com",
        "CdenHartog@Ambaflex.com",
        "cdenhartog@m-asset.nl",
    ]
    department_filter = "Alles,Department,Inside Sales EMEA,Field Service AMCAS Employees,Warehouse RO,R&D,Customer Service EMEA,Customer Service AMCAS,Product Engineering NL,Warehouse US,Field Service EMEA,Sales EMEA,Inside Sales AMCAS,Sales APAC,Sales AMCAS,Inside Sales APAC,Global Order Management,Global Customer Service Department,Assembly US,Warehouse NL,Field Service APAC Employees,Warehouse CN,Finance & Internal Control,Finance,Customer Service APAC,Spare Parts Warehouse US Shift 1 Employees,Construction US,Sector Operations,Order Management RO,Hands-on Development,Global Sales Operations,Finance RO,Finance CN,Assembly RO,Sales Project Management,Order Management US,IT & PMO,Finance US,Construction NL,Production US,Materials Management,IBD Packaging,HR RO,HR NL,Construction CN,Assembly NL,Assembly CN,Tactical Purchasing US Employees,Replenishing RO Employees,Order Management NL,None,Medanva Employees,Injection Moulding US,IBD AccuVeyor,Global Operations Support,Global Application Engineering,Replenishing CN,Quality Engineering US,Production RO,Production CN,Order Engineering CN Employees,MES,Internal Control,Injection Moulding NL Team Lead,IBD Logistics,HR US,HR CN Employees,HR & Organizational Development,Global Customer Service,Facility US Employees,Facility RO,Construction RO,AmbaFlex Integration Service,Support Desk,Spare Parts Warehouse NL Employees,Sector Sales,Sector Innovation,Sector Business Administration,Replenishing NL Employees,Quality Engineering CN Employees,Project Management,Production NL,Product Engineering,OMD - PDM,Marketing,Logiflexx Conveyor Solutions,Interim Inkoper,Innovation Development,Industry Sub Chains,IT CN Employees (Kunshan),IBD B&C,IBD Airport,Global Sales & Support Team,Facility CN Employees (Kunshan),Facility,Corporate,Coating RO"
    lost_and_found = ["avanriesen@ambaflex.com"]

    for i in range(len(lost_and_found)):
        lost_and_found[i] = lost_and_found[i].lower()

    print("get data from sharefile (Citrix)")
    sharefile_header = loginCitrixShareFile()
    if sharefile_header == None:
        exit()
    print("get data from Entra ID")
    token = get_graph_access_token()
    sharefile_entra_group_users = get_all_users_of_group(token, "GS_NL_CTX_ShareFile")
    all_entra_users, filtered_users = get_all_Entra_Users(
        token, department_filter, lost_and_found
    )
    sharefile_employees, count = get_full_list_of_employees(sharefile_header)
    output = ""
    # add id to filtered_users for each sharefile employee with the same email
    email_to_id = {
        sharefile_employee["Email"].lower(): sharefile_employee["Id"]
        for sharefile_employee in sharefile_employees
    }
    name_to_id = {
        sharefile_employee["Email"].lower(): sharefile_employee["Name"]
        for sharefile_employee in sharefile_employees
    }
    isDisabled_to_id = {
        sharefile_employee["Email"].lower(): sharefile_employee["IsDisabled"]
        for sharefile_employee in sharefile_employees
    }

    print("Start filtering")

    for user in filtered_users:
        user["sharefile_id"] = email_to_id.get(user["userPrincipalName"].lower(), None)
        user["Name"] = name_to_id.get(user["userPrincipalName"].lower(), None)
        user["IsSharefileEnabled"] = isDisabled_to_id.get(
            user["userPrincipalName"].lower(), -1
        )
        # print(str(user["Name"]) + str(user["IsSharefileEnabled"]))

        user["HomeFolderSizeMB"] = get_home_folder_by_user_id(
            user["sharefile_id"], sharefile_header
        )
    for user in all_entra_users:
        user["sharefile_id"] = email_to_id.get(user["userPrincipalName"].lower(), None)
        user["Name"] = name_to_id.get(user["userPrincipalName"].lower(), None)

    # Add "InShareFileGroup" column to filtered_users
    sharefile_entra_group_user_ids = {
        user["id"] for user in sharefile_entra_group_users
    }
    for user in filtered_users:
        user["InShareFileGroup"] = user["id"] in sharefile_entra_group_user_ids

    department_employees_with_sharefile = {
        employee["userPrincipalName"]
        for employee in filtered_users
        if employee["sharefile_id"] is not None
    }

    filtered_users.sort(
        reverse=True,
        key=lambda x: (
            x["sharefile_id"] is not None,
            x["InShareFileGroup"],
            x["accountEnabled"],
            not (x["IsSharefileEnabled"]),
            x["HomeFolderSizeMB"],
        ),
    )
    print("Check Departments")

    departments = []
    department_counts = {}

    for user in all_entra_users:
        department = str(user.get("department"))
        sharefile_id = user.get("sharefile_id")
        if sharefile_id:
            if department:
                if department not in department_counts:
                    department_counts[department] = {
                        "description": department,
                        "user_count": 0,
                        "name": "",
                    }
                department_counts[department]["user_count"] += 1
                department_counts[department]["name"] = (
                    department_counts[department]["name"] + ", " + user["Name"]
                )

    # Convert the dictionary to a list of department details
    departments = list(department_counts.values())
    departments.sort(
        reverse=True,
        key=lambda x: (x["user_count"], x["description"]),
    )

    # Call the function to create the Excel file
    create_excel_with_users(filtered_users, department_filter)
    add_department_usages_to_excel(filtered_users, sharefile_header, department_filter)
    add_totals_to_excel(
        filtered_users,
        sharefile_entra_group_users,
        department_filter,
        sharefile_employees,
    )

    add_open_departments_to_excel(departments)
