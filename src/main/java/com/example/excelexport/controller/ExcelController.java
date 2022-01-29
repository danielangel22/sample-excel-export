package com.example.excelexport.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.example.excelexport.User;
import com.example.excelexport.util.ExcelGenerator;

@RestController
public class ExcelController {

	@Autowired
	private ExcelGenerator excelGenerator;

	private List<User> users = new ArrayList<>();

	@GetMapping(value = "/export/excel", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void exportToExcel(HttpServletResponse response) throws IOException {
		String headerValue = "attachment; filename=users.xlsx";
		response.setHeader(HttpHeaders.CONTENT_DISPOSITION, headerValue);
		if (users.isEmpty())
			generateUserList();
		excelGenerator.export(response, users);
	}

	public List<User> generateUserList() {
		var response = new ArrayList<User>();
		var random = new Random();
		for (var i = 0; i < 100; i++) {
			var b = random.nextInt() + "-User";
			response.add(new User(i, b, b, "name", true));
		}
		users = response;
		return response;
	}

}
