package execPy;

import opXML.AnalyXML;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

public class ExecPy implements Runnable{

    private String docFileName;
    private String dirPath;

    public ExecPy(String dirPath, String docFileName) {
        this.docFileName = docFileName;
        this.dirPath = dirPath;
    }

    @Override
    public void run() {

        String classPath = String.valueOf(ExecPy.class.getResource(""));
        classPath = classPath.substring(classPath.indexOf("/"));
        System.out.println("classPath : " + classPath);
        String checkPy = /*PyDirPath + */classPath + "check.py";
        String modifyPy = /*PyDirPath + */classPath + "modify.py";
        String DataDirPath = dirPath + docFileName + "/";
        System.out.println("dirPath: " + dirPath);
        System.out.println("docFileName: " + docFileName);

        try {

            System.out.println("check start");
            Process checkProcess = Runtime.getRuntime().exec("python3.5" + " " + checkPy + " " + classPath/*PyDirPath*/ + " " + DataDirPath);
            String line;

            BufferedReader inputStream = new BufferedReader(new InputStreamReader(checkProcess.getInputStream()));
            while ((line = inputStream.readLine()) != null) {
                System.out.println(line);
            }
            inputStream.close();

            BufferedReader errorStream = new BufferedReader(new InputStreamReader(checkProcess.getErrorStream()));
            while ((line = errorStream.readLine()) != null) {
                System.out.println(line);
            }
            errorStream.close();

            checkProcess.waitFor();//等待check程序执行完毕
            System.out.println("check end");

            System.out.println("modify start");
            Process modifyProcess = Runtime.getRuntime().exec("python3.5" + " " + modifyPy + " " + DataDirPath);

            inputStream = new BufferedReader(new InputStreamReader(modifyProcess.getInputStream()));
            while ((line = inputStream.readLine()) != null) {
                System.out.println(line);
            }
            inputStream.close();

            errorStream = new BufferedReader(new InputStreamReader(modifyProcess.getErrorStream()));
            while ((line = errorStream.readLine()) != null) {
                System.out.println(line);
            }
            errorStream.close();

            modifyProcess.waitFor();//等待modify程序执行完毕
            System.out.println("modify end");

            try {
                System.out.println("add comment start");
                new AnalyXML(DataDirPath).run();
                System.out.println("add comment end");
            }
            catch (Exception e) {
                System.err.println("添加批注出现错误");
                e.printStackTrace();
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        String classPath = String.valueOf(ExecPy.class.getResource(""));
        System.out.println("classPath : " + classPath.substring(classPath.indexOf("/") + 1));
    }
}
