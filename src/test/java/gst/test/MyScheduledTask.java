package gst.test;

import gstdemo.ExcelRead;
public class MyScheduledTask implements Runnable {
    @Override
    public void run() {
        // Place your task code here
        System.out.println("Scheduled task is running...");
        ExcelRead read = new ExcelRead();
        read.execute(); // Assuming FilteredDataTest has a method named execute()
        
    }
}