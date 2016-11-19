<%@ page import="java.io.BufferedReader" %>
<%@ page import="java.io.FileReader" %>
<%@ page import="java.io.File" %>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>论文格式校正</title>
</head>
<body>
    <%
        try {
            String absolutePath = "C:/Users/Billy/Documents/GitHub/OfficialProgram/Paper/data/";
            String dirName = request.getParameter("dirName");
            File txtFile = new File(absolutePath + dirName + "/check_out.txt");
            FileReader fileReader = new FileReader(txtFile);
            BufferedReader bf = new BufferedReader(fileReader);
            String readLine = bf.readLine();
            while (readLine != null) {
                out.print(readLine);
                out.print("<br>");
                readLine = bf.readLine();
            }

            out.println("<br>");
            out.println("<br>");
            out.print("<a href='checkWordResult.jsp?dirName=");
            out.print(dirName);
            out.print("'>");
            out.print("下载Word文档");
            out.print("</a>");
        }
        catch (Exception e) {
            response.sendRedirect("showErrorInfo.jsp");
        }
    %>
    <br>
    <br>
    <%--<a href="checkWordResult.jsp?dirName=<%=dirName%>">下载Word文档</a>--%>
</body>
</html>
