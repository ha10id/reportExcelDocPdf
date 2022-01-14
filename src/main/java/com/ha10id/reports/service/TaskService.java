/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.ha10id.reports.service;

import com.ha10id.reports.entity.Task;
import com.ha10id.reports.repository.TaskRepository;
import java.util.List;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

/**
 *
 * @author ha10id
 */
@Service
public class TaskService {
    
    @Autowired
    private TaskRepository taskRepository;
    
    public List<Task> getAll() {
        return taskRepository.findAll(Sort.by(Sort.Order.asc("date"),
                                      Sort.Order.desc("priorityId")));
    }

    public Task save(Task task) {
        return taskRepository.save(task);
    }

    public void delete(int id) {
        taskRepository.deleteById(id);
    }
}
