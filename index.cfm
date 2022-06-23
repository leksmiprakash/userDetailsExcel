<cfset showMessage = false>

<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Document</title>
        <link href="./css/bootstrap.min.css" rel="stylesheet">
    </head>
    <body>
    <cfoutput>
        <cfif StructKeyExists(session, "errors")>
            <div id="message" class="alert alert-success" role="alert">
                #session.errors#
            </div>
        </cfif>
        <cfif StructKeyExists(session, "success")>
           <cfset session.success= true>
        <cfelse>
            <cfset session.success= false>
        </cfif>
        <section>
        
            <div class="d-flex  mt-5 p-3">
                <div class="col-3">
                    <a href="upload/Plain_Template.xlsx" class="btn btn-info" download>Plane Template</a>
                </div>
                <div class="col-3">
                    <a href="component/userDetails.cfc?method=excelDownload" class="btn btn-info">Template With Data</a>
                </div>
                <div class="col">
                    <form name="form" method="post" action="component/userDetails.cfc?method=excelUpload" enctype="multipart/form-data"> 
                        <div class="row">
                            <div class="col-2">
                                <label for="inputFile" class="btn btn-secondary">
                                    <input type="file" name="inputFile" id="inputFile" class="d-none">
                                    Browse
                                </label>
                            </div>
                            <div class="col-4">
                                <button type="submit" class="btn btn-success" >Upload</button>
                            </div>
                        </div>
                    </form>
                </div>
            </div>
            
        </section>
            <cfset UserObj=CreateObject("component","component.userDetails")/>
            <cfset users=UserObj.displayUserData()/>
            <div>
                <table class="table table-bordered bg-light m-5">
                    <thead>
                        <tr>
                            <th>First Name</th>
                            <th>Last Name</th>
                            <th>Address</th>
                            <th>Email</th>
                            <th>Phone</th>
                            <th>DOB</th>
                            <th>Role</th>
                        </tr>
                    </thead>
                    <tbody>
                        <cfloop QUERY="#users#">
                            <tr>
                                <td>#users.firstName#</td>
                                <td>#lastName#</td>
                                <td>#address#</td>
                                <td>#email#</td>
                                <td>#phone#</td>
                                <td>#DateFormat(dob)#</td>
                                <td>#roleofUser#</td>
                            </tr>
                        </cfloop>
                    </tbody>
                </table>
            </div>  
            <script src="js/jquery-3.6.0.min.js"></script> 
            <script src="js/bootstrap.min.js"></script>
        <script>
            const redirectPage = () => {
                if((#session.success#==true ) && (#url?.success#)){
                    window.location = "autoDownloadExcel.cfm";
                }
                <cfset session.success = false>
            }
            setTimeout(redirectPage(), 3000);
        </script>
        
            </div>
        </div>
    </cfoutput>
    </body> 
    
</html>