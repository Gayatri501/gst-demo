package gst.test;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
public class MainTest {
    public static void main(String[] args) {
    	
        ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

        MyScheduledTask myTask = new MyScheduledTask();

        
        long intervalInSeconds = 15;
        scheduler.scheduleAtFixedRate(myTask, 0L, intervalInSeconds, TimeUnit.SECONDS);
        
        //scheduler.shutdown();
    }
}
