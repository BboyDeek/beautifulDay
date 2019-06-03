package com.example.demo.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;
/**
 *首页
 *@param:
 *@return:
 *@Author:zym
 *@Date:2019/6/3
 */
@Controller
public class IndexController {
	@RequestMapping("/")
	public String fileTest(HttpServletRequest request, HttpSession session) {
		return "redirect:fileTest.html";
	}
}