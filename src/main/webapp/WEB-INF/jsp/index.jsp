<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html>
<html lang="en">
	<head>
		<meta charset="UTF-8">
		<title>Document</title>
	</head>
	<body>
    <input type="hidden" id="sbmId" value="${sbmId}">
	<div id="app" style="width:100%; height:100%;"></div>

    <script type="text/javascript" src="${frontName}/sbm.js?${uuid}"></script>
    <script>
        var sbm = new SBM('#app')
    </script>
</body>
</html>