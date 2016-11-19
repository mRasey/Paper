<%@ page contentType="text/html;charset=gbk" language="java" %>
<html lang="zh" class="no-js">
    <style>
        input[type="image"] {
            background: url("res/image/upload.png") no-repeat;
            width: 128px;
            height: 128px;
            border: none;
        }

        input {
            border: none;
        }

    </style>
    <head>
        <title>首页</title>
        <link rel="stylesheet" type="text/css" href="css/css-title/normalize.css" />
        <link rel="stylesheet" type="text/css" href="fonts/font-awesome-4.2.0/css/font-awesome.min.css" />
        <link rel="stylesheet" type="text/css" href="css/css-title/demo.css" />
        <link rel="stylesheet" type="text/css" href="css/css-title/component.css" />

        <%--<meta charset="UTF-8">--%>
        <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <link rel="stylesheet" type="text/css" href="css/css-upload/normalize.css" />
        <%--<link rel="stylesheet" type="text/css" href="css/css-upload/default.css" />--%>
        <link rel="stylesheet" type="text/css" href="css/css-upload/component.css" />
        <link rel="stylesheet" type="text/css" href="css/css-button/button.css" />
        <link rel="stylesheet" href="dist/styles/Vidage.css" />
        <script>(function(e,t,n){
            var r=e.querySelectorAll("html")[0];
            r.className=r.className.replace(/(^|\s)no-js(\s|$)/,"$1js$2")})(document,window,0);
        </script>


    </head>
    <script>
        function checkFileType() {
            var fileName = document.getElementById("file").value;
            var preFix = fileName.substring(fileName.lastIndexOf("\\") + 1, fileName.length - 5);
            var postfix = fileName.substring(fileName.length - 5, fileName.length);
            for(var i = 0; i < preFix.length; i++) {
                if(preFix.charAt(i) < '0' || preFix.charAt(i) > '9') {
                    alert("不合法的文件名，请重试");
                    return false;
                }
            }
            if(postfix == ".docx") {
                return true;
            }
            else {
                alert("请输入有效的.docx文件");
                return false;
            }
        }
    </script>
    <body>
        <div class="container">
            <header class="codrops-header">
                <h1>北航毕设论文格式在线校正系统<span>（测试版）</span></h1>
                <p style="color:white">注意：请将论文命名为：学号.docx，不要带有中文，docx文件不要带有批注</p>
            </header>
            <div class="content">
                <div class="box">
                    <form onsubmit="return checkFileType()" method="post" action="UploadFile.jsp" enctype="multipart/form-data">
                    <input type="file" name="file" id="file" class="inputfile inputfile-3" data-multiple-caption="{count} files selected" multiple />
                    <label for="file">
                    <span>Choose a file&hellip;</span></label>
                    <br><br>
                    <%--<input type="file" name="file" id="file"><br />--%>
                    <%--<input type="submit" class="button" value="提交">--%>
                    <input name="imgbtn" type="image" src="res/image/upload.png">
                    <div style="font-size: small">点击上传</div>
                    <%--<a class="button" href="javascript:;" onclick="next();">提交</a>--%>
                    </form>
                </div>
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
    <script src="js/custom-file-input.js"></script>
</html>
