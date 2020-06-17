package com.globalhitss.scanPC;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ScanPcApplication implements CommandLineRunner {

	@Autowired
	ScanPC scanPC;

	public static void main(String[] args) {

		SpringApplication.run(ScanPcApplication.class, args);

		
	}

	@Override
	public void run(String... args) throws Exception {
		// Start Scan
		this.scanPC.scan();
	}

}
