package com.vkb.print.servlet;

import com.vkb.print.DocCreator;
import org.eclipse.jetty.server.Server;
import org.eclipse.jetty.servlet.ServletContextHandler;
import org.eclipse.jetty.servlet.ServletHolder;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

public class VKBPrintServlet extends HttpServlet {

    @Override
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {

        String locations = req.getParameter("LOCATIONS");
        String [] locationArr = locations.split(",");

        Map<String, String> paragraphs = new LinkedHashMap<>();

        for(String location : locationArr) {
            String monthlyList =  req.getParameter(location);
            String infrequentList =  req.getParameter("INFREQUENT " + location);

            if(monthlyList != null && ! monthlyList.isEmpty()) {
                paragraphs.put(location, monthlyList);
            }

            if(infrequentList != null && ! infrequentList.isEmpty()) {
                paragraphs.put("INFREQUENT " + location, infrequentList);
            }

            System.out.println("Mthly " + monthlyList);
            System.out.println("Infrequent " + infrequentList);
        }

        String fileName = "C:/Temp/list.docx";
        writeToFile(paragraphs, fileName);

//        printFromFile(fileName);
    }

    private void printFromFile(String fileName) throws IOException {
        File ff = new File(fileName);
        Desktop desktop = Desktop.getDesktop();
        desktop.print(ff);
    }

    private void writeToFile(Map<String, String> paragraphs, String fileName) throws FileNotFoundException {
        DocCreator.createDocFile(fileName, paragraphs);
    }

    public static void main(String[] args) throws Exception {
        String port = "8000";
        Server server = new Server(Integer.valueOf(port));

        ServletContextHandler context = new ServletContextHandler(ServletContextHandler.SESSIONS);
        context.setContextPath("/");
        server.setHandler(context);
        context.addServlet(new ServletHolder(new VKBPrintServlet()),"/printGroceryList");

        server.start();
        server.join();
    }
}
