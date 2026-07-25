
import java.io.*;
import java.nio.file.*;
import java.util.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import com.microsoft.playwright.*;
import com.microsoft.playwright.assertions.LocatorAssertions;
import com.microsoft.playwright.options.AriaRole;
import static com.microsoft.playwright.assertions.PlaywrightAssertions.*;

public class App {
	public static XWPFDocument document;
	public static List<XWPFParagraph> paragraphs;
	public static Map<String, String> documentValues = new HashMap<>();
	public static XSSFWorkbook containers;
	public static String[][] containerValues;
	public static int containerCount = 0;

	public static void main(String[] args) throws InterruptedException, IOException {
		readWord();
		readExcel();
		//System.exit(0);
		Playwright playwright = Playwright.create();
		Browser browser = playwright.chromium().launch(new BrowserType.LaunchOptions().setHeadless(false));
		Page page = browser.newPage();
		page.navigate("https://www.shipmentlink.com/tam1/jsp/TAM1_Login.jsp");
		// This is a terrible security method. However, it is mainly used to prevent
		// credential stealing from quick glances from bystanders.
		page.fill("//input[@name = \"id\"]",
				new String(Base64.getDecoder().decode("input base64 encode")));
		page.fill("//input[@name = \"password\"]", new String(Base64.getDecoder().decode("input base64 encode")));
		assertThat(page.getByText("Partner with EVERGREEN LINE"))
				.isVisible(new LocatorAssertions.IsVisibleOptions().setTimeout(100000000));
		System.out.println("OK Status");
		page.click("//*[@id=\"idx_al_in_one\"]/div/div/div[3]/ul/li[1]/a");
		page.getByRole(AriaRole.BUTTON, new Page.GetByRoleOptions().setName("Agree")).click();
		assertThat(page.getByText("Create/Modify new B/L Instruction", new Page.GetByTextOptions().setExact(true)))
				.isVisible(new LocatorAssertions.IsVisibleOptions().setTimeout(100000000));

		/* WORD DOCUMENT */
		page.fill("//*[@id=\"SHIPPER_TYPE\"]", documentValues.get("SHIPPER"));
		page.fill("//*[@id=\"CONSIGNEE_TYPE\"]", documentValues.get("CONSIGNEE"));
		page.fill("//*[@id=\"NOTIFY_PARTY_1_TYPE\"]", documentValues.get("NOTIFY PARTY"));
		page.fill("//*[@id=\"NOTIFY_PARTY_2_TYPE\"]", documentValues.get("2ND NOTIFY PARTY"));
		page.fill("//*[@id=\"EXPORT_REFERENCE_TYPE\"]", documentValues.get("EXPORT REFERENCE"));
		page.fill("//*[@name=\"shippingMarks\"]", documentValues.get("SHIPPING MARK"));
		page.fill("//*[@name=\"AES_NO\"]", documentValues.get("XTN"));
		page.fill("//*[@name=\"cargoDesp\"]", documentValues.get("DESCRIPTION"));
		page.fill("//*[@name=\"remarks\"]", documentValues.get("COMMENTS"));
		/* DOCUMENT CONTACT */
		page.fill("//*[@id=\"EBI1_CONTACT_PERSON\"]", "NAME");
		page.fill("//*[@id=\"EBI1_TEL_ZIP\"]", "TEL NUMBER PART 1");
		page.fill("//*[@id=\"EBI1_TEL\"]", "TEL NUMBER PART 2");
		page.fill("//*[@id=\"EBI1_TEL_EXT\"]", "TEL NUMBER PART 3");
		page.click("//*[@id=\"EBI1_MORE_EMAIL\"]");
		page.fill("//*[@id=\"EBI1_NTFY_EMAIL\"]", "EMAIL 1");
		page.fill("//*[@id=\"EBI1_MORE_EMAIL_1\"]", "EMAIL 2");

		/* EXCEL DOCUMENT / CONTAINER INFORMATION */
		for (int i = 0; i < containerCount; i++) {
			page.locator("//*[@name=\"containerNo\"]").nth(i).fill(containerValues[i][0]);
			page.locator("//*[@name=\"grossWeight\"]").nth(i).fill(containerValues[i][3]);
			page.locator("//*[@name=\"ebi_cnt_quantity\"]").nth(i).fill(containerValues[i][2]);
			page.locator("//*[@name=\"quantityUnit\"]").nth(i).selectOption("OPTION");
			page.locator("//*[@name=\"sealNo1\"]").nth(i).fill(containerValues[i][1]);
//			page.locator("//*[@name=\"sealNo2\"]").nth(i).fill("45");
//			page.locator("//*[@name=\"sealNo3\"]").nth(i).fill("987");
//			page.locator("//*[@name=\"sealNo4\"]").nth(i).fill("4");
//			page.locator("//*[@name=\"sealNo5\"]").nth(i).fill("325");
			if (i != containerCount - 1) {
				page.getByRole(AriaRole.BUTTON, new Page.GetByRoleOptions().setName("Add Container")).click();
			}
		}
		System.out.println("Done");

	}

	public static void readWord() throws IOException, InterruptedException {
		document = new XWPFDocument(Files.newInputStream(Paths
				.get("INPUT FILE ADDRESS")));
		paragraphs = document.getParagraphs();
		documentValues.put("SHIPPING MARK", "N/A");
		documentValues.put("2ND NOTIFY PARTY", "");
		documentValues.put("COMMENTS", "");

		for (int i = 0; i < paragraphs.size(); i++) {
			String p = paragraphs.get(i).getText();
			if (p.length() > 28 && p.substring(0, 27).equals("BILL OF LADING INSTRUCTIONS")) {
				sameLiner("EXPORT REFERENCE", p);
				continue;
			}
			if (p.length() > 7 && p.substring(0, 7).equals("BOOKING")) {
				sameLiner("BOOKING", p);
				continue;
			}
			if (p.length() > 6 && p.substring(0, 7).equals("SHIPPER")) {
				i = concatenate("SHIPPER", i);
				continue;
			}
			if (p.length() > 9 && p.substring(0, 9).equals("CONSIGNEE")) {
				i = concatenate("CONSIGNEE", i);
				continue;
			}
			if (p.length() > 11 && p.substring(0, 12).equals("NOTIFY PARTY")) {
				i = concatenate("NOTIFY PARTY", i);
				continue;
			}
			if (p.length() > 15 && p.substring(0, 16).equals("2ND NOTIFY PARTY")) {
				i = concatenate("2ND NOTIFY PARTY", i);
				continue;
			}
			if (p.length() > 11 && p.contains("ITN")) {
				sameLiner("ITN", p);
				continue;
			}
			if (p.length() > 9 && p.contains("XTN")) {
				sameLiner("XTN", p);
				continue;
			}
			if (p.length() > 11 && p.substring(0, 11).equals("DESCRIPTION")) {
				i = concatenate("DESCRIPTION", p.substring(13) + " \n", i);
				continue;
			}
			if (p.length() > 12 && p.substring(0, 13).equals("SHIPPING MARK")) {
				sameLiner("SHIPPING MARK", p);
				continue;
			}
			if (p.length() > 4 && p.substring(0, 5).equals("LC NO")) {
				documentValues.put("COMMENTS", documentValues.get("COMMENTS") + p + "\n");
				continue;
			}
			if (p.length() > 12 && p.substring(0, 13).equals("DATE OF ISSUE")) {
				documentValues.put("COMMENTS", documentValues.get("COMMENTS") + p + "\n");
				continue;
			}
			if (p.length() > 12 && p.contains("UNDER BONDED WAREHOUSE")) {
				documentValues.put("COMMENTS", documentValues.get("COMMENTS") + p + "\n");
				continue;
			}
			if (p.length() > 14 && p.substring(0, 15).equals("FREIGHT PREPAID")) {
				documentValues.put("COMMENTS", documentValues.get("COMMENTS")
						+ "FREIGHT PREPAID, DTHC IS PREPAID \nBILL OF LADING MUST SHOW NAME, ADDRESS, TELEPHONE AND FAX NUMBER OF LOCAL SHIPPING AGENT \n");
				continue;
			}
			if (p.length() > 3 && p.substring(0, 2).equals("**")) {
				documentValues.put("COMMENTS", documentValues.get("COMMENTS") + p + "\n");
			}
		}
		if (documentValues.get("COMMENTS").equals("")
				|| !documentValues.get("COMMENTS").contains("BILL OF LADING MUST")) {
			documentValues.put("COMMENTS", documentValues.get("COMMENTS")
					+ "BILL OF LADING MUST SHOW NAME, ADDRESS, TELEPHONE AND FAX NUMBER OF LOCAL SHIPPING AGENT\n");
		}
		System.out.println(documentValues.get("EXPORT REFERENCE"));
		System.out.println(documentValues.get("BOOKING"));
		System.out.println(documentValues.get("SHIPPER"));
		System.out.println(documentValues.get("CONSIGNEE"));
		System.out.println(documentValues.get("NOTIFY PARTY"));
		System.out.println(documentValues.get("2ND NOTIFY PARTY"));
		System.out.println(documentValues.get("SHIPPING MARK"));
		System.out.println(documentValues.get("ITN"));
		System.out.println(documentValues.get("XTN"));
		System.out.println(documentValues.get("DESCRIPTION"));
		System.out.println(documentValues.get("COMMENTS"));

//		for(XWPFParagraph p : paragraphs) {
//			System.out.println(p.getText());
//			Thread.sleep(1000);
//		}
		document.close();
	}

	public static void readExcel() throws IOException, InterruptedException {
		DataFormatter dataFormatter = new DataFormatter();
		containers = new XSSFWorkbook(
				"INPUT FILE ADDRESS");
		Sheet sheet = containers.getSheetAt(0);

		for (int i = 6; i < 34; i++) {
			Row row = sheet.getRow(i);
			Cell cell = row.getCell(0);
			if (dataFormatter.formatCellValue(cell).equals("")) {
				break;
			}
			int value = Integer.parseInt(dataFormatter.formatCellValue(cell));
			if (containerCount < value) {
				containerCount = value;
			}
		}
		containerValues = new String[containerCount][4];
		for (int i = 0; i < containerCount; i++) {
			Row row = sheet.getRow(i + 6);
			// System.out.println(dataFormatter.formatCellValue(row.getCell(0)));
			for (int j = 0; j < 4; j++) {
				containerValues[i][j] = dataFormatter.formatCellValue(row.getCell(j + 1));
			}
		}
//		for (int i = 0; i < containerCount; i++) {
//			for (int j = 0; j < 4; j++) {
//				System.out.println(containerValues[i][j]);
//			}
//			System.out.println();
//		}

	}

	public static int concatenate(String bound, int i) {
		i++;
		String content = "";
		while (!paragraphs.get(i).getText().equals("")) {
			content += paragraphs.get(i).getText();
			content += "\n";
			i++;
		}
		documentValues.put(bound, content);
		return i;
	}

	public static int concatenate(String bound, String fragment, int i) {
		i++;
		String content = fragment;
		while (!paragraphs.get(i).getText().equals("")) {
			content += paragraphs.get(i).getText();
			content += "\n";
			i++;
		}
		documentValues.put(bound, content);
		return i;
	}

	public static void sameLiner(String bound, String p) {
		String[] temp = p.split(" ");
		documentValues.put(bound, temp[temp.length - 1] + "\n");
	}

	public static String splitter(String p) {
		String[] temp = p.split(" ");
		return temp[temp.length - 1] + "\n";
	}
}
