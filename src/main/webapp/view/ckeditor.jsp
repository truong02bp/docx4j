<%@ taglib prefix="form" uri="http://www.springframework.org/tags/form" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <script src="https://code.jquery.com/jquery-1.9.1.min.js"></script>
    <script src="<c:url value='/template/ckeditor/ckeditor.js'/>"></script>
    <title>Document</title>
</head>
<body>
<form id="formSubmit" method="post">
    <label for="content">Content</label>
    <textarea id="content" rows="100" cols="150"></textarea>
    <button type="button" id="submit">Submit</button>
</form>
</body>
<script>
    var editor="";
    $(document).ready(function () {
        editor = CKEDITOR.replace("content");
    });
    $("#submit").click(function (e){
        e.preventDefault();
        var data = {};
        data["content"] = editor.getData();
        $.ajax({
                url: "/add",
                method : "POST",
                contentType: 'application/json',
                data: JSON.stringify(data),
                success: function (result) {
                    window.location.href = "http://localhost:8081/docx?id="+result.id;
                },
                error: function (error) {
                    console.log("errors");
                }
            }
        )
    })
</script>
</html>
