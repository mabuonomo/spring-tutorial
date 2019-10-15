package it.mabuonomo.javatoexcel.controller;

import java.util.ArrayList;
import java.util.List;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import it.mabuonomo.javatoexcel.component.ExcelView;
import it.mabuonomo.javatoexcel.model.Report;

@RestController
@RequestMapping("/report")
public class ReportController {

  @RequestMapping("/example")
  public ModelAndView getCallReportXls() {

    List<Report> reports = new ArrayList<Report>();

    Report rep1 = new Report();
    rep1.setName("rep1");
    rep1.setValue(10.0);

    Report rep2 = new Report();
    rep2.setName("rep2");
    rep2.setValue(22.22);

    reports.add(rep1);
    reports.add(rep2);

    return new ModelAndView(new ExcelView(), "reports", reports);
  }
}
