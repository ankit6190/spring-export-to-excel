package com.example.demo.controller;


import com.example.demo.config.ExportExcelView;
import com.example.demo.domain.Student;
import org.jxls.template.SimpleExporter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@RestController
public class StudentController {

    @Autowired
    ExportExcelView exportExcelView;

    @RequestMapping(value = "/students/", method = RequestMethod.GET)
    public List<Student> getStudents(){
        List<Student> studentList = new ArrayList<Student>();
        studentList.add(Student.builder().firstName("Test").lastName("Test").build());
        studentList.add(Student.builder().firstName("Test1").lastName("Test1").build());
        studentList.add(Student.builder().firstName("Test1").lastName("Test1").build());
        studentList.add(Student.builder().firstName("Test1").lastName("Test1").build());
        return studentList;
    }

    @RequestMapping(value = "/students/export", method = RequestMethod.GET)
    public void export(HttpServletResponse response) {
        List<Student> studentList = new ArrayList<Student>();
        studentList.add(Student.builder().firstName("Test").lastName("Test").build());
        studentList.add(Student.builder().firstName("Test1").lastName("Test1").build());
        studentList.add(Student.builder().firstName("Test1").lastName("Test1").build());
        studentList.add(Student.builder().firstName("Test1").lastName("Test1").build());

        List<String> headers = Arrays.asList("First Name", "Last Name");
        try {
            response.addHeader("Content-disposition", "attachment; filename=People.xlsx");
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            new SimpleExporter().gridExport(headers, studentList, "firstName, lastName, ", response.getOutputStream());
            response.flushBuffer();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @RequestMapping(value = "/students/export1", method = RequestMethod.GET)
    public ModelAndView export1(ModelAndView modelAndView, HttpServletRequest request, HttpServletResponse response) throws Exception {
        List<Student> studentList = new ArrayList<Student>();
        studentList.add(Student.builder().firstName("Test").lastName("Test").build());
        studentList.add(Student.builder().firstName("Test1").lastName("Test1").build());
        studentList.add(Student.builder().firstName("Test1").lastName("Test1").build());
        studentList.add(Student.builder().firstName("Test1").lastName("Test1").build());

        String[] header = { "First Name", "Last Name"};

        // we have to pass bean field name which we need to show in excel
        String[] beanMapColumn = { "firstName", "lastName"};

        modelAndView.addObject("dataList", studentList);
        modelAndView.addObject("headerList", header);
        modelAndView.addObject("beanMapColumn", beanMapColumn);

        exportExcelView.setFileName("Test");
        exportExcelView.setBeanName("com.example.demo.domain.Student");
        exportExcelView.setSheetName("Student List");
        exportExcelView.buildExcelDocument(modelAndView.getModel(), null, request, response);
        modelAndView.setView(exportExcelView);
        return modelAndView;
    }
}
