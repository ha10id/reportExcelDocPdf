/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.ha10id.reports.controller;

import com.ha10id.reports.entity.Task;
import com.ha10id.reports.service.DocToPdfService;
import com.ha10id.reports.service.TaskService;
import java.io.IOException;
import java.util.Date;
import java.util.List;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 *
 * @author ha10id
 */
@Controller
public class TaskController {

    @Autowired
    private TaskService taskService;
    @Autowired
    private DocToPdfService dtpService;

    @GetMapping("/")
    public String getAll(Model model) {
        List<Task> taskList = taskService.getAll();
        model.addAttribute("taskList", taskList);
        model.addAttribute("taskSize", taskList.size());
        return "index";
    }

    @RequestMapping("/delete/{id}")
    public String deleteTask(@PathVariable int id) {
        taskService.delete(id);
        return "redirect:/";
    }

    @PostMapping("/add")
    public String addTask(@ModelAttribute Task task) {
        taskService.save(task);
        return "redirect:/";
    }

    @PostMapping("/adds")
    public String addsTask() {
        Task task = new Task(1, 1, "Test task 1", new Date());
        taskService.save(task);
        for (int i = 2; i < 10; i++) {
            task = new Task(i, 1, "Test task " + i, new Date());
            taskService.save(task);
        }
        return "redirect:/";
    }

    @PostMapping("/doc-to-pdf")
    public String docToPdf() throws IOException, Exception {
        dtpService.convert("./test.xlsx", "./toPDF.pdf");
        return "redirect:/";
    }
}
