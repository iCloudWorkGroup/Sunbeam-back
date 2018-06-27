<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html lang="en">
	<head>
		<meta charset="UTF-8">
		<title>Document</title>
	</head>
	<body>
    <input type="hidden" id="sbmId" value="${sbmId}">
	<div id="app" style="width:100%; height:100%;"></div>

    <script type="text/javascript" src="${frontName}/dist/sbm.js"></script>
    <script>
        var sbm = new SBM('#app')
    </script>
</body>
</html>