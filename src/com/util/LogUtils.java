package com.util;

import javax.swing.JTextPane;

public interface LogUtils {
	public static void setTextToLogScreen(JTextPane textPane_log,org.slf4j.Logger logger ,String log )
	{
		logger.info(log);
		String someHtmlMessage = "<html><b style='color:blue;'>"+log+"</b><html>";
		textPane_log.setText(textPane_log.getText()+"\n"+log);	
	}

}
