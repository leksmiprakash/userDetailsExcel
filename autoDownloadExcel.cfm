<cfheader name="Content-Disposition" value="attachment;filename=Results.xls">
<cfcontent type="application/vnd.msexcel" variable="#SpreadSheetReadBinary(session.spreadsheet)#" reset="false">