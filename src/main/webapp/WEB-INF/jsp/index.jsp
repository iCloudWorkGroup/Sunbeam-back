<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>Document</title>
	<style>
		.table{
			height:100%
		}
	</style>
</head>
<body>
	<input type="hidden" id="auth_key" value="${sbmId}">
	<div id="app" style="width:100%; height:100%;">
		<div class="tools"></div>
		<div class="table"></div>
	</div>
	<script type="text/javascript" src="${frontName}/sunbeam.js?${sbmId}"></script>
   <script>
        var sbm = new Sunbeam({
            root: '.table',
            toolbar: '.tools'
        })
		sbm.load().then(() => {
			
		})
    </script>
</body>
</html>