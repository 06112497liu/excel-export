package com.lwb.excel.export;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@MapperScan(basePackages = {"com.lwb.excel.export.dao"})
@SpringBootApplication
public class ExcelExportApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExcelExportApplication.class, args);
	}

}
