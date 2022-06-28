<cfcomponent displayname="userdata" hint="Data from user side">

    <cffunction name="displayUserData" access="public" returnType="any" output="false">
        <cfquery name = "getUsers">
            SELECT usertable.*, GROUP_CONCAT(roles.roleTitle SEPARATOR ', ')  AS roleofUser
            FROM(
                (
                    usertable JOIN userroles ON userroles.userId = usertable.userId
                )
                JOIN roles ON userroles.roleId = roles.roleId
            )
            GROUP BY userroles.userId;
        </cfquery>
        <cfreturn getUsers> 
    </cffunction>
   
    <cffunction  name="excelUpload" returntype ="struct" access="remote" output="false">
        <cfset local.thisDir = expandPath(".")>
        <cfset local.errors = "">
        <cfset local.success = false>
        <cfset local.errorRows = 2>
        <cfset local.errorFreeRows = 0>
        <cffile action="upload" destination="#local.thisDir#/upload" filefield="inputFile" result="upload" nameconflict="makeunique">
        <cfif upload.fileWasSaved>
            <cfset local.savedFile = upload.serverDirectory & "\" & upload.serverFile>
            <cfif isSpreadsheetFile(local.savedFile)>
                <cfspreadsheet action="read" src="#local.savedFile#" query="data" headerrow="1">
                <cfset local.columnNames = 'First Name,Last Name,Address,Email,Phone,DOB,Role,Result'>
                <cfif data.recordCount is 1>
                    <cfset local.errors = " Empty spreadsheet.<br>">
                    <cfset session.errors = errors>
                <cfelse>
                    <cfset spreadsheet = spreadsheetNew("Users") />
                    <cfset SpreadsheetSetActiveSheet(spreadsheet, "Users")/>
                    <cfloop from="1" to="#listLen(local.columnNames)#" index="i">
                        <cfset SpreadsheetSetCellValue(spreadsheet, listGetAt(local.columnNames, i) ,  1, i) />
                    </cfloop>
                    <cfquery name="getRoles" returntype="array">
                        select *
                        from roles;
                    </cfquery>
                    <cfset local.rowErrorMsg = "">
                    <cfset local.rolesArray = arrayNew(1)>
                    <cfloop array="#getRoles#" item="roleFromQuery">
                        <cfset arrayAppend(local.rolesArray, roleFromQuery.roleTitle)>
                    </cfloop>
                    <cfset local.totalValidRows = 1>
                    <cfloop index="rows" from="2" to="#data.recordCount#">
                        
                        
                        <cfset local.totalValidRows = local.totalValidRows+1>
                    </cfloop>
                    <cfloop index="row" from="2" to="#data.recordCount#">
                        <cfset local.rowError = false>
                        <cfset local.rowErrorMsg = "">
                        <cfset local.emptyRow = 1>
                        <cfloop index="emptyCheckCol" from="1" to="#listLen(local.columnNames)-1#">
                            <cfif len(data[listGetAt(local.columnNames, emptyCheckCol)][row]) EQ 0>
                                <cfset local.emptyRow = local.emptyRow+1>
                            </cfif>
                        </cfloop>
                        <cfif local.emptyRow GTE 7>
                            <cfcontinue>
                        </cfif>
                        <cfloop index="col" from="1" to="#listLen(local.columnNames)#">
                            <cfif listGetAt(local.columnNames, col) != 'Result'>
                                <cfif len(data[listGetAt(local.columnNames, col)][row]) GT 0>
                                    <cfif listGetAt(local.columnNames, col) == 'First Name'>
                                        <cfif !isValid("regex", data[listGetAt(local.columnNames, col)][row], "^[a-zA-Z ]*$")>
                                            <cfset local.rowError = true>
                                            <cfif len(local.rowErrorMsg) GT 0>
                                                <cfset local.rowErrorMsg = listAppend(local.rowErrorMsg, 'First Name can only have alphabets and space')>
                                            <cfelse>
                                                <cfset local.rowErrorMsg = 'First Name can only have alphabets and space'>
                                            </cfif>
                                        </cfif>
                                    </cfif>
                                    <cfif listGetAt(local.columnNames, col) == 'Last Name'>
                                        <cfif !isValid("regex", data[listGetAt(local.columnNames, col)][row], "^[a-zA-Z ]*$")>
                                            <cfset local.rowError = true>
                                            <cfif len(local.rowErrorMsg) GT 0>
                                                <cfset local.rowErrorMsg = listAppend(local.rowErrorMsg, 'Last Name can only have alphabets and space')>
                                            <cfelse>
                                                <cfset local.rowErrorMsg = 'Last Name can only have alphabets and space'>
                                            </cfif>
                                        </cfif>
                                    </cfif>
                                    <cfif listGetAt(local.columnNames, col) == 'Email'>
                                        <cfif !isValid("email", data[listGetAt(local.columnNames, col)][row])>
                                            <cfset local.rowError = true>
                                            <cfif len(local.rowErrorMsg) GT 0>
                                                <cfset local.rowErrorMsg = listAppend(local.rowErrorMsg, 'Enter a valid Email')>
                                            <cfelse>
                                                <cfset local.rowErrorMsg = 'Enter a valid Email'>
                                            </cfif>
                                        </cfif>
                                    </cfif>
                                    <cfif listGetAt(local.columnNames, col) == 'Phone'>
                                        <cfif !isValid("regex", data[listGetAt(local.columnNames, col)][row],"^(\+\d{1,2}\s?)?1?\-?\.?\s?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$")>
                                            <cfset local.rowError = true>
                                            <cfif len(local.rowErrorMsg) GT 0>
                                                <cfset local.rowErrorMsg = listAppend(local.rowErrorMsg, 'Enter a valid Phone Number')>
                                            <cfelse>
                                                <cfset local.rowErrorMsg = 'Enter a valid Phone Number'>
                                            </cfif>
                                        </cfif>
                                    </cfif>
                                    <cfif listGetAt(local.columnNames, col) == 'Role'>
                                        <cfset local.allRoleExist = true>
                                        <cfset local.roleIds = arrayNew(1)>
                                        <cfloop list="#data[listGetAt(local.columnNames, col)][row]#" item="roleFromRow">
                                            <cfif !arrayContains(local.rolesArray, roleFromRow)>
                                                <cfset local.allRoleExist = false>
                                            <cfelse>
                                                <cfloop array="#getRoles#" item="roleFromQuery">
                                                    <cfif  roleFromQuery.roleTitle EQ roleFromRow>
                                                        <cfset arrayAppend(local.roleIds,  roleFromQuery.roleId)>
                                                    </cfif>
                                                </cfloop>
                                            </cfif>
                                        </cfloop>
                                        <cfif !local.allRoleExist>
                                            <cfset local.rowError = true>
                                            <cfif len(local.rowErrorMsg) GT 0>
                                                <cfset local.rowErrorMsg = listAppend(local.rowErrorMsg, 'Roles are not valid')>
                                            <cfelse>
                                                <cfset local.rowErrorMsg = 'Roles are not valid'>
                                            </cfif>
                                        </cfif>
                                    </cfif>
                                <cfelse>
                                    <cfset local.rowError = true>
                                    <cfif len(local.rowErrorMsg) GT 0>
                                        <cfset local.rowErrorMsg = listAppend(local.rowErrorMsg, '#listGetAt(local.columnNames, col)# is missing')>
                                    <cfelse>
                                        <cfset local.rowErrorMsg = '#listGetAt(local.columnNames, col)# is missing'>
                                    </cfif>
                                </cfif>
                            </cfif>
                        </cfloop>
                        <cfif local.rowError>
                            <cfloop index="colIndex" from="1" to="#listLen(local.columnNames)-1#">
                                <cfset SpreadsheetSetCellValue(spreadsheet, data[listGetAt(local.columnNames, colIndex)][row] , local.errorRows, colIndex) />
                            </cfloop>
                            <cfset SpreadsheetSetCellValue(spreadsheet, local.rowErrorMsg , local.errorRows, 8) />
                            <cfset  local.errorRows =  local.errorRows+1>
                        <cfelse>
                            <cfquery name="userExist">
                                select userId
                                from usertable
                                where email = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Email'][row]#">
                            </cfquery>
                            <cfset local.queryExcecuteSucceded = true>
                            <cftry>
                                <cfif queryRecordCount(userExist) GT 0>
                                    <cfquery name="updateUser" result="updatedRow">
                                        UPDATE usertable 
                                        SET 
                                            firstName = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['First Name'][row]#">, 
                                            lastName = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Last Name'][row]#">, 
                                            address = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Address'][row]#">, 
                                            email = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Email'][row]#">, 
                                            phone = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Phone'][row]#">, 
                                            dob = <cfqueryparam cfsqltype="cf_sql_date" value="#data['DOB'][row]#">
                                        WHERE usertable.userId = <cfqueryparam cfsqltype="cf_sql_integer" value="#userExist.userId#">
                                    </cfquery>
                                    <cfquery name="removeRolesOfUser">
                                        DELETE FROM `userroles` WHERE `userId` = <cfqueryparam cfsqltype="cf_sql_integer" value="#userExist.userId#">;
                                    </cfquery>
                                    <cfloop array="#local.roleIds#" item="roleId"> 
                                        <cfquery name="addUserRoles">
                                            INSERT INTO userroles (
                                                userId,roleId) 
                                            VALUES (
                                                <cfqueryparam cfsqltype="cf_sql_integer" value="#userExist.userId#">, 
                                                <cfqueryparam cfsqltype="cf_sql_integer" value="#roleId#">)
                                        </cfquery>
                                    </cfloop>
                                <cfelse>
                                    <cfquery name="addUser" result="addedUser">
                                        INSERT INTO usertable (
                                            firstName, 
                                            lastName, 
                                            address, 
                                            email, 
                                            phone,
                                            dob) 
                                        VALUES (
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['First Name'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Last Name'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Address'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Email'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Phone'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_timestamp" value="#data['DOB'][row]#">)
                                    </cfquery>
                                    <cfloop array="#local.roleIds#" item="roleId"> 
                                        <cfquery name="addUserRoles">
                                            INSERT INTO userroles (
                                                userId,roleId) 
                                            VALUES (
                                                <cfqueryparam cfsqltype="cf_sql_integer" value="#addedUser.generatedKey#">, 
                                                <cfqueryparam cfsqltype="cf_sql_integer" value="#roleId#">)
                                        </cfquery>
                                    </cfloop>
                                </cfif>
                            <cfcatch type="any">
                                <cfset local.queryExcecuteSucceded = false>
                                <cfloop index="colIndex" from="1" to="#listLen(local.columnNames)-1#">
                                    <cfset SpreadsheetSetCellValue(spreadsheet, data[listGetAt(local.columnNames, colIndex)][row] , local.errorRows, colIndex) />
                                </cfloop>
                                <cfset SpreadsheetSetCellValue(spreadsheet, '#cfcatch.message#' , local.errorRows, 8) />
                                <cfset  local.errorRows =  local.errorRows+1>
                            </cfcatch>
                            </cftry>
                            <cfif local.queryExcecuteSucceded>
                                <cfloop index="colIndex" from="1" to="#listLen(local.columnNames)-1#">
                                    <cfset SpreadsheetSetCellValue(spreadsheet, data[listGetAt(local.columnNames, colIndex)][row] , local.totalValidRows-local.errorFreeRows, colIndex) />
                                </cfloop>
                                <cfif queryRecordCount(userExist) GT 0>
                                    <cfset SpreadsheetSetCellValue(spreadsheet, 'Updated' ,   local.totalValidRows-local.errorFreeRows, 8) />
                                <cfelse>
                                    <cfset SpreadsheetSetCellValue(spreadsheet, 'Added' ,   local.totalValidRows-local.errorFreeRows, 8) />
                                </cfif>
                                <cfset  local.errorFreeRows =  local.errorFreeRows  +1>
                            </cfif>
                        </cfif>
                    </cfloop>
                    <cfset local.success = true>
                    <cfset session.success = true>
                    <cfset session.spreadsheet = spreadsheet>
                    <cflocation  url="../index.cfm?success=true" addtoken="false">
                </cfif>
            <cfelse>
                <cfset local.errors = "File format dosen't match.<br>">
                <cfset session.errors = errors>
            </cfif>
        <cffile  action="delete" file="local.savedFile">
        <cfelse>
            <cfset local.errors = "Something went wrong.<br>">
            <cfset session.errors = errors>	
        </cfif>
        <cfset returnData = structNew()>
        <cfset returnData["success"] = local.success>
        <cfset returnData["errors"] = local.errors>
        <cfset returnData["savedFile"] = local.savedFile>
        <cfif local.success>
            <cfset returnData["spreadsheet"] = spreadsheet>
        </cfif>
        <cfreturn returnData>
    </cffunction>

     <cffunction  name="excelDownload" access="remote">
        <cfset getAllUsers = displayUserData()>
        <cfset spreadsheet = spreadsheetNew("UsersList") />
        <cfset SpreadsheetSetActiveSheet(spreadsheet, "UsersList")/>
        <cfset SpreadsheetSetCellValue(spreadsheet, "First Name",  1, 1) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Last Name", 1, 2)/>
        <cfset SpreadsheetSetCellValue(spreadsheet, "Address", 1, 3) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Email", 1, 4) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Phone", 1, 5) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "DOB", 1, 6) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Role", 1, 7) />
        <cfloop index="row" from="1" to="#getAllUsers.recordCount#">
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['firstName'][row],  row+1, 1) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['lastName'][row], row+1, 2)/>
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['address'][row], row+1, 3) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['email'][row], row+1, 4) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['phone'][row], row+1, 5) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['dob'][row], row+1, 6) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['roleofUser'][row], row+1, 7) />
        </cfloop>
        <cfheader name="Content-Disposition" value="inline; filename=users.xls">
        <cfcontent type="application/vnd.msexcel" variable="#SpreadSheetReadBinary(spreadsheet)#">
    </cffunction>

</cfcomponent>