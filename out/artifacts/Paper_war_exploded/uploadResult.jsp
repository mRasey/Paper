<%@ page import="execPy.ExecPy" %>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%
    String dirPath = request.getParameter("dirPath");
    dirPath = dirPath.replaceAll("\\\\", "/");
    String fileName = request.getParameter("fileName");
    fileName = fileName.replaceAll("\\\\", "/");
    System.err.println(dirPath);
    new Thread(new ExecPy(dirPath, fileName)).start();
%>
<html>
<head>
    <title>上传结果</title>
    <link rel="stylesheet" type="text/css" href="css/css-title/normalize.css" />
    <link rel="stylesheet" type="text/css" href="fonts/font-awesome-4.2.0/css/font-awesome.min.css" />
    <link rel="stylesheet" type="text/css" href="css/css-title/demo.css" />
    <link rel="stylesheet" type="text/css" href="css/css-title/component.css" />
    <link rel="stylesheet" type="text/css" href="css/css-button/button.css" />
    <%--<link rel="stylesheet" type="text/css" href="css/css-loading/bootstrap.min.css">--%>
    <%--<link rel="stylesheet" type="text/css" href="css/css-loading/htmleaf-demo.css">--%>
    <link rel="stylesheet" type="text/css" href="css/css-loading/loading.css">
    <link rel="stylesheet" href="dist/styles/Vidage.css" />


</head>
<%--<script type="text/javascript" src="js/jquery-2.1.1.min.js"></script>--%>
<%--<script type="text/javascript" src="js/jquery-1.5.1.js"></script>--%>
<script type="text/javascript" src="js/jquery-3.1.0.js"></script>
<script type="text/javascript" charset="gbk">
    function loopingAsk() {
        document.getElementById("upSuc").style.display = "none";
        document.getElementById("loading").style.display = "";
        setInterval(execute, 100);
    }
    function execute() {
        var i = 0;
//        alert(fileName);
//        while(true) {
            $.ajax({
                type: "GET",
                dataType: "json",
                url: "Scan?dirPath=" + "<%=dirPath%>" + "&fileName=" + "<%=fileName%>"
            }).done(function (data) {
                var ifFind = eval(data);
                if (data.toString() == "true") {
//                    alert(data);
                    self.location = 'result.jsp?dirPath=' + "<%=dirPath%>" + "&fileName=" + "<%=fileName%>";
                    clearInterval(execute);
                }
            }).fail(function () {
                alert("处理失败");
                self.location = "showErrorInfo.jsp";
            });
//        }
    }
</script>
<body>
    <div class="container">
        <header class="codrops-header">
            <h1>北航毕设论文格式在线校正系统<span>（测试版）</span></h1>
        </header>
        <div class="htmleaf-container" id="loading" style="display: none">
            <div class="demo" >
                <div class="container">
                    <div class="row">
                        <div class="col-md-12">
                            <div class="loader">
                                <div class="loading-1"></div>
                                <div class="loading-2">处理中...</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <div align="center" id="upSuc">
            <p style="color:white">上传成功</p>
            <br><br>
            <a class="button" id="start" onclick="loopingAsk()">开始处理</a>
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
