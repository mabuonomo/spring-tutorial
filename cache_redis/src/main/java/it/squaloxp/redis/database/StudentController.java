package it.squaloxp.redis.database;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
@RequestMapping(path = "/database")
public class StudentController {

    @Autowired
    StudentRepository studentRepository;

    @GetMapping(path = "/save")
    public @ResponseBody Student save() {
        Student student = new Student("Eng2015001", "John Doe", Student.Gender.MALE, 1);
        return studentRepository.save(student);
    }

    @GetMapping(path = "/get")
    public @ResponseBody Student get() {
        return studentRepository.findById("Eng2015001").get();
    }

    @GetMapping(path = "/add")
    public @ResponseBody Student add() {
        Student randStudent = new Student(Integer.toString(new Random().nextInt()), "Rand Houston", Student.Gender.MALE, 3);
        return studentRepository.save(randStudent);
    }

    @GetMapping(path = "/all")
    public @ResponseBody List<Student> all() {
        List<Student> students = new ArrayList<>();
        studentRepository.findAll().forEach(students::add);
        return students;
    }

    @GetMapping(path = "/clear")
    public @ResponseBody List<Student> clear() {
        studentRepository.deleteAll();

        List<Student> students = new ArrayList<>();
        studentRepository.findAll().forEach(students::add);
        return students;
    }
}
