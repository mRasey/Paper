package servlets;

import net.sf.json.JSONObject;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;

public class Scan extends HttpServlet{

//    private HttpServletRequest request;
//    private HttpServletResponse response;

    public Scan() {
        super();
    }


//    public Scan(HttpServletRequest request, HttpServletResponse response) {
//        this.request = request;
//        this.response = response;
//    }

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        String fileName = req.getParameter("fileName");
        String dirPath = req.getParameter("dirPath");
//            String dirPath = "";
        System.err.println(fileName);
        System.out.println(dirPath);
        try {
//            JSONObject jsonObject = new JSONObject();
            //req.getRequestDispatcher(req.getSession().getServletContext().getRealPath("/") + "result.jsp").forward(req,resp);
            File dir = new File(dirPath + fileName);
            if (dir.listFiles() != null) {
                for (File file : dir.listFiles()) {
                    if (file.getName().equals("resultWithComments.docx")) {
//                        System.err.println(req.getSession().getServletContext().getRealPath("/") + "/result.jsp");
//                        jsonObject.put("ifFind", "true");
                        resp.getWriter().print("true");
                        resp.getWriter().close();
//                            req.getRequestDispatcher(req.getSession().getServletContext().getRealPath("/") + "result.jsp").forward(req,resp);
//                            resp.sendRedirect(req.getSession().getServletContext().getRealPath("/") + "result.jsp?fileName=" + fileName);
                        System.err.println("end");
                        return;
                    }
                }
            }
//            jsonObject.put("ifFind", "false");
            resp.getWriter().print("false");
            resp.getWriter().close();
        }catch (Exception e) {
            e.printStackTrace();
//            try {
//                resp.sendRedirect("showErrorInfo.jsp");
//            } catch (IOException e1) {
//                e1.printStackTrace();
//            }
        }
    }
}
