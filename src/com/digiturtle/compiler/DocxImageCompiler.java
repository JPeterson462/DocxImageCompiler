package com.digiturtle.compiler;

import java.io.File;
import java.util.ArrayList;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileSystemView;

public class DocxImageCompiler {

	private static int getIndex(String path, String prefix) {
		return Integer.parseInt(path.substring(prefix.length(), path.lastIndexOf('.')).trim());
	}

	public static void main(String[] args) throws Exception {
		JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		int returnValue = jfc.showOpenDialog(null);
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			File selectedFile = jfc.getSelectedFile();
			System.out.println(selectedFile.getAbsolutePath());

			String prefix = JOptionPane.showInputDialog("Image Filter");
			File[] possibleImages = selectedFile.listFiles();
			
			DocxBuilder builder = new DocxBuilder("test1.docx");
			ArrayList<File> files = new ArrayList<>();
			for (File possibleImage : possibleImages) {
				if (!possibleImage.getName().contains("ignore") && possibleImage.getName().startsWith(prefix)) {
					files.add(possibleImage);
				}
			}
			files.sort((f1, f2) -> Integer.compare(getIndex(f1.getName(), prefix), getIndex(f2.getName(), prefix)));
			for (File file : files) {
				System.out.println(file.getName());
				builder.addImageToWordDocument(file);
			}
			builder.close();			
		}
		else {
			throw new Exception("Directory required!");
		}
	}

	

}
