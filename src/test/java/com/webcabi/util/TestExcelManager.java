package com.webcabi.util;

import static org.junit.Assert.fail;

import java.io.File;

import org.junit.AfterClass;
import org.junit.After;
import org.junit.BeforeClass;
import org.junit.Before;
import org.junit.Test;

public class TestExcelManager {
	@BeforeClass
	public static void setUpBeforeClass() throws Exception {
	}
	
	@AfterClass
	public static void tearDownAfterClass() throws Exception {
	}
	
	@Before
	public void setUp() throws Exception {
	}
	
	@After
	public void tearDown() throws Exception {
	}
	
	@Test
	public void test() {
		ExcelManager manager = new ExcelManager();
		manager.open(new File(""));
	}
}
