<%@ page contentType="text/html;charset=gbk" language="java" pageEncoding="gbk" autoFlush="true" %>
<html>
<head>
    <title>正在处理</title>
</head>
<body>

</body>
<%@ page import="java.io.*,java.util.*, javax.servlet.*" %>
<%@ page import="javax.servlet.http.*" %>
<%@ page import="org.apache.commons.fileupload.*" %>
<%@ page import="org.apache.commons.fileupload.disk.*" %>
<%@ page import="org.apache.commons.fileupload.servlet.*" %>
<%@ page import="org.apache.commons.io.output.*" %>
<%@ page import="opXML.ExtractXML" %>

<%
    try {
        System.err.println("in uploadFile");
        String fileCode = (String) System.getProperties().get("file.encoding");
        File file;
        int maxFileSize = 5000 * 1024;
        int maxMemSize = 5000 * 1024;
        ServletContext context = pageContext.getServletContext();
        String filePath = request.getSession().getServletContext().getRealPath("/") + "data/";
//        System.err.println(dirPath);
//        String filePath = "C:\\Users\\Billy\\Documents\\GitHub\\OfficialProgram\\Paper\\data\\";
        // 验证上传内容了类型
        String contentType = request.getContentType();
        if ((contentType.indexOf("multipart/form-data") >= 0)) {

            DiskFileItemFactory factory = new DiskFileItemFactory();
            // 设置内存中存储文件的最大值
            factory.setSizeThreshold(maxMemSize);
            // 本地存储的数据大于 maxMemSize.
            factory.setRepository(new File("c:\\temp"));

            // 创建一个新的文件上传处理程序
            ServletFileUpload upload = new ServletFileUpload(factory);
            // 设置最大上传的文件大小
            upload.setSizeMax(maxFileSize);
            try {
                // 解析获取的文件
                List fileItems = upload.parseRequest(request);

                // 处理上传的文件
                Iterator i = fileItems.iterator();

                out.println("<html>");
                out.println("<head>");
                out.println("<title>JSP File upload</title>");
                out.println("</head>");
                out.println("<body>");
                while (i.hasNext()) {
                    FileItem fi = (FileItem) i.next();
                    if (!fi.isFormField()) {
                        // 获取上传文件的参数
                        String fieldName = fi.getFieldName();
                        String fileName = fi.getName();//上传文件名
                        fileName = new String(fileName.getBytes(fileCode), fileCode);//按照系统默认编码重新读取文件名
                        boolean isInMemory = fi.isInMemory();
                        long sizeInBytes = fi.getSize();
                        // 写入文件
                        String name;//上传文件的短名称
                        if (fileName.lastIndexOf("\\") >= 0) {
                            name = fileName.substring(fileName.lastIndexOf("\\") + 1);
                            File dir = new File(filePath + name);
                            if(dir.exists())
                                ExtractXML.deleteAllDir(filePath, name);
                            new File(filePath + name).mkdirs();
                            file = new File(filePath + name + "/", "origin.docx");//上传文件命名为origin
                        } else {
                            name = fileName.substring(fileName.lastIndexOf("\\") + 1);
                            File dir = new File(filePath + name);
                            if(dir.exists())
                                ExtractXML.deleteAllDir(filePath, name);
                            new File(filePath + name).mkdirs();
                            file = new File(filePath + name + "/", "origin.docx");//上传文件命名为origin
                        }
                        fi.write(file);
//                    new File(filePath + name + "\\", "check_out.txt").createNewFile();
//                    new File(filePath + name + "\\", "check_out1.txt").createNewFile();
                        out.println("Uploaded Filename: " + filePath + fileName + "<br>");
                        response.sendRedirect("uploadResult.jsp?dirPath=" + filePath + "&fileName=" + name);//上传成功，跳转到上传结果界面
                    }
                }
                out.println("</body>");
                out.println("</html>");
            } catch (Exception ex) {
                response.sendRedirect("showErrorInfo.jsp");
                System.out.println(ex);
            }
        } else {
            out.println("<html>");
            out.println("<head>");
            out.println("<title>Servlet upload</title>");
            out.println("</head>");
            out.println("<body>");
            out.println("<p>No file uploaded</p>");
            out.println("</body>");
            out.println("</html>");
        }
    }
    catch (Exception e) {
        response.sendRedirect("showErrorInfo.jsp");
    }
%>
</html>
