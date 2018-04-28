package com.acmr.excel.util;

public class UploadThread implements Runnable{
	private String partFileName;
	private byte[] fileByte;

	@Override
	public void run() {
		ExcelUtil.write(partFileName, fileByte);
	}

	public UploadThread(String partFileName, byte[] fileByte) {
		this.partFileName = partFileName;
		this.fileByte = fileByte;
	}
}
