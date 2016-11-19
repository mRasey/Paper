<%@ page import="java.io.File" %>
<%@ page import="java.io.FileReader" %>
<%@ page import="java.io.BufferedReader" %>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>处理成功</title>
    <link rel="stylesheet" type="text/css" href="css/css-title/normalize.css" />
    <link rel="stylesheet" type="text/css" href="fonts/font-awesome-4.2.0/css/font-awesome.min.css" />
    <link rel="stylesheet" type="text/css" href="css/css-title/demo.css" />
    <link rel="stylesheet" type="text/css" href="css/css-title/component.css" />
    <link rel="stylesheet" type="text/css" href="css/css-button/button.css" />
    <link rel="stylesheet" href="dist/styles/Vidage.css" />

</head>
<body>
    <%
        String fileName = request.getParameter("fileName");
        String dirPath = request.getParameter("dirPath");
        if(!new File(dirPath + fileName + "/result.docx").exists()
                || !new File(dirPath + fileName + "/resultWithComments.docx").exists())
            response.sendRedirect("showErrorInfo.jsp");

//        String txtPath = dirPath + fileName + "/" + "check_out.txt";
//        String wordPath = dirPath + fileName + "/" + "result.docx";
    %>
    <%--<jsp:include page="<%=txtPath%>" flush="true"/>--%>
    <head>
        <title>论文格式校正</title>
    </head>
    <div class="container">
        <header class="codrops-header">
            <h1>北航毕设论文格式在线校正系统<span>（测试版）</span></h1>
            <%--<p>测试版</p>--%>
        </header>
        <div align="center">
            <%--查看结果<br><br>--%>
            <%--<a href="checkTXTInfo.jsp?dirName=<%=dirName%>">查看错误信息</a><br>--%>
            <a href="checkWordResult.jsp?dirPath=<%=dirPath%>&fileName=<%=fileName%>">修改过的Word文档</a><br><br>
            <a href="checkWordWithComments.jsp?dirPath=<%=dirPath%>&fileName=<%=fileName%>">带有批注的原Word文档</a><br><br><br>
            <a class="button" href="index.jsp">回到首页</a>
            <br><br>
            <p style="color:white">
                Microsoft Office 2007及以上的版本可在菜单栏的审阅栏目下找到“比较”来比对修改前后的不同
            </p>
        </div>
    </div>

    <div class="Vidage">
        <div class="Vidage__image"></div>

        <video id="VidageVideo" class="Vidage__video" preload="metadata" loop autoplay muted>
            <source src="videos/bg.webm" type="video/webm">
            <source src="videos/bg.mp4" type="video/mp4">
        </video>

        <div class="Vidage__backdrop"></div>
    </div>


    <!-- Vidage init -->
    <script src="dist/scripts/Vidage.min.js"></script>
    <script>
        new Vidage('#VidageVideo');
    </script>
</body>
</html>
