Public Sub Main()
	Set obj = createobject("ADODB.Connection")						'Creating an ADODB Connection Object
	Set obj1 = createobject("ADODB.RecordSet")						'Creating an ADODB Recordset Object
	Dim dbquery														'Declaring a database query variable bquery 
	Dbquery = "Select acctno from dbo.acct where name = 'Harsh'"	'Creating a query 
	obj.Open "Provider=SQLQLEDB;Server=.\SQLEXPRESS;UserId=test;Password=P@123;Database =AUTODB"    'Opening a Connection    
	obj1.Open dbquery,obj			'Executing the query using recordset  
	val1 = obj1.fields.item(0)		'Will return field value 
	msgbox val1                     'Displaying value of the field item 0 i.e. column 1
	obj.close                       'Closing the connection object
	obj1.close                      'Closing the connection object
	Set obj1 = Nothing				'Releasing Recordset object
	Set obj = Nothing					'Releasing Connection object
End Sub