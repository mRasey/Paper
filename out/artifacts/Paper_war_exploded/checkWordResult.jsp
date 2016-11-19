<%@page import="java.util.*"%>
<%@page import="java.io.*"%>
<%@page import="java.net.*"%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>论文格式校正</title>
</head>
<body>
    <%
        try {
            String dirPath = request.getParameter("dirPath");
            String fileName = request.getParameter("fileName");
            response.setContentType("application/msword");
            response.setHeader("Content-disposition", "attachment; filename=result.docx");
            BufferedInputStream bis = null;
            BufferedOutputStream bos = null;
            try {
                bis = new BufferedInputStream(new FileInputStream(new File(dirPath + fileName + "/result.docx")));
                bos = new BufferedOutputStream(response.getOutputStream());
                byte[] buff = new byte[2048];
                int bytesRead;
                while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                    bos.write(buff, 0, bytesRead);
                }
            } catch (final IOException e) {
                System.out.println("出现IOException." + e);
            } finally {
                if (bis != null)
                    bis.close();
                if (bos != null)
                    bos.close();
            }
        }
        catch (Exception e) {
            response.sendRedirect("showErrorInfo.jsp");
        }
    %>
</body>
</html>
