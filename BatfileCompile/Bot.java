import java.awt.*;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.*;

import static java.lang.Character.isUpperCase;
import static java.lang.Character.toUpperCase;

public class Bot {
    public static void main(String[] args) throws Exception {
        File file = new File("C:\\Users\\Mingcan\\Desktop\\BatteryTester\\Testingsoftware\\LiBatTest.exe");
        Desktop.getDesktop().open(file);
        Thread.sleep(2000);
        Robot robot = new Robot();
        File folder = new File("C:\\Users\\Mingcan\\Desktop\\DCIR");
        File[] list = folder.listFiles();
        Boolean unfolded = false;
        for(File current : list) {
            if(current.getName().endsWith(".cel")) {
                robot.mouseMove(130, 75); //move mouse to open file button
                Thread.sleep(500);
                robot.mousePress(InputEvent.BUTTON1_MASK); //click the open file button
                robot.mouseRelease(InputEvent.BUTTON1_MASK);
                Thread.sleep(1000);
                enterString(current.getAbsolutePath(), robot); //enter the absolute path
                robot.keyPress(KeyEvent.VK_ENTER); //click open
                robot.keyRelease(KeyEvent.VK_ENTER);
                Thread.sleep(500);
                robot.mouseMove(600, 200);
                robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
                robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
                Thread.sleep(100);
                if (!unfolded) {
                    robot.keyPress(KeyEvent.VK_DOWN); //move highlight to unfold
                    robot.keyRelease(KeyEvent.VK_DOWN);
                    Thread.sleep(100);
                    robot.keyPress(KeyEvent.VK_ENTER);
                    robot.keyRelease(KeyEvent.VK_ENTER);
                    Thread.sleep(100);
                    unfolded = true;

                }
                robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
                robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
                for (int i = 0; i < 3; i++) {  //move highlight to Generate text
                    robot.keyPress(KeyEvent.VK_DOWN);
                    robot.keyRelease(KeyEvent.VK_DOWN);
                }

                Thread.sleep(200);
                for (int i = 0; i < 2; i++) {  //press enter twice
                    robot.keyPress(KeyEvent.VK_ENTER);
                    robot.keyRelease(KeyEvent.VK_ENTER);
                    Thread.sleep(100);
                }
                Thread.sleep(500);
//            robot.mouseMove(200,750);
//            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
//            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
//            robot.mouseMove(180,720);
//            robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
//            robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
//            Thread.sleep(500);
//            enterString(current.getName().replace(".cel",""),robot);
//            robot.keyPress(KeyEvent.VK_ENTER);
//            robot.keyRelease(KeyEvent.VK_ENTER);
                robot.mouseMove(1000, 168);
                robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
                robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                Thread.sleep(1000);
            }
        }
    }



    public static void enterString(String text, Robot robot) throws InterruptedException {
        //robot won't play nice with symbols so I have to resolve some of them individually   s\Mingcan\Desktop\BatFiles\_220817_box1_group5_001_3_2.cel
        
        for(int i = 0 ; i < text.length() ; i++) {
            if(text.charAt(i)=='\\') {
                robot.keyPress(KeyEvent.VK_BACK_SLASH);
                robot.keyRelease(KeyEvent.VK_BACK_SLASH);
            }
            else if(text.charAt(i)=='_') {
                robot.keyPress(KeyEvent.VK_SHIFT);
                robot.keyPress(toUpperCase('-'));
                robot.keyRelease(toUpperCase('-'));
                robot.keyRelease(KeyEvent.VK_SHIFT); 
            }


            else if(text.charAt(i)==':') {
                robot.keyPress(KeyEvent.VK_SHIFT);
                robot.keyPress(KeyEvent.VK_SEMICOLON);
                robot.keyRelease(KeyEvent.VK_SEMICOLON);
                robot.keyRelease(KeyEvent.VK_SHIFT);
            
            }else{
                if (isUpperCase(text.charAt(i)))
                    robot.keyPress(KeyEvent.VK_SHIFT);
                char u=Character.toUpperCase(text.charAt(i));
                System.out.print(u);
                robot.keyPress(Character.toUpperCase(text.charAt(i)));
                robot.keyRelease(Character.toUpperCase(text.charAt(i)));
                robot.keyRelease(KeyEvent.VK_SHIFT);
            }
            Thread.sleep(10);
        }
    }
}
