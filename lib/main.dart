// ignore_for_file: lines_longer_than_80_chars

import "dart:developer";
import "dart:io";

import "package:document_file_save_plus/document_file_save_plus.dart";
import "package:excel/excel.dart";
import "package:flutter/foundation.dart";
import "package:flutter/material.dart";
import "package:universal_html/html.dart" as uni;

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: "Generate & Download Excel Demo",
      theme: themeData(Brightness.light),
      darkTheme: themeData(Brightness.dark),
      home: const MyHomePage(),
      debugShowCheckedModeBanner: false,
    );
  }

  ThemeData themeData(Brightness brightness) {
    return ThemeData(
      useMaterial3: true,
      brightness: brightness,
      colorSchemeSeed: Colors.blue,
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key});

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text("Generate & Download Excel Demo"),
      ),
      body: SafeArea(
        child: Center(
          child: Padding(
            padding: const EdgeInsets.all(16.0),
            child: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              children: <Widget>[
                const Text(
                  "This is a simple app that helps you generate and download Excel files. You can use it on Android, iOS, and the web.",
                ),
                const SizedBox(height: 16),
                const Text(
                  "On Android, if you click Generate Excel, the file will be saved in your Downloads folder.",
                ),
                const SizedBox(height: 16),
                const Text(
                  "For iOS, when you press Generate Excel, a menu will pop up with choices like copying the file or saving it.",
                ),
                const SizedBox(height: 16),
                const Text(
                  "If you're using the web version, hitting Generate Excel can either ask you to download the file or save it straight to your Downloads folder.",
                ),
                const SizedBox(height: 32),
                ElevatedButton(
                  onPressed: generateExcelFile,
                  child: const Text("Generate Excel"),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }

  Future<void> generateExcelFile() async {
    final Excel excel = Excel.createExcel();
    final Sheet? sheet = excel.sheets[excel.getDefaultSheet() ?? ""];
    final List<String> rowHeader = <String>[
      "Column #1",
      "Column #2",
      "Column #3",
      "Column #4",
      "Column #5",
    ];

    for (int i = 0; i < rowHeader.length; i++) {
      final CellIndex cellIndex0 = indexByColumnRow(cIndex: i, rIndex: 0);

      sheet?.cell(cellIndex0).value = rowHeader[i];
    }

    for (int i = 1; i <= 10; i++) {
      final CellIndex cellIndex0 = indexByColumnRow(cIndex: 0, rIndex: i);
      final CellIndex cellIndex1 = indexByColumnRow(cIndex: 1, rIndex: i);
      final CellIndex cellIndex2 = indexByColumnRow(cIndex: 2, rIndex: i);
      final CellIndex cellIndex3 = indexByColumnRow(cIndex: 3, rIndex: i);
      final CellIndex cellIndex4 = indexByColumnRow(cIndex: 4, rIndex: i);

      sheet?.cell(cellIndex0).value = "Row #$i";
      sheet?.cell(cellIndex1).value = "Row #$i";
      sheet?.cell(cellIndex2).value = "Row #$i";
      sheet?.cell(cellIndex3).value = "Row #$i";
      sheet?.cell(cellIndex4).value = "Row #$i";
    }

    final String currentSheetName = excel.getDefaultSheet() ?? "";
    const String updatedSheetName = "Sheet 1";
    const String downloadFileName = "sample.xlsx";
    const String mimeType = "application/vnd.ms-excel";
    excel.rename(currentSheetName, updatedSheetName);

    final List<int> intList = excel.save(fileName: downloadFileName) ?? <int>[];
    final Uint8List uint8 = Uint8List.fromList(intList);

    if (kIsWeb) {
      final uni.Blob blob = uni.Blob(uint8, mimeType, "native");
      final String href = uni.Url.createObjectUrlFromBlob(blob);
      uni.AnchorElement(
        href: href,
      )
        ..setAttribute("download", downloadFileName)
        ..click();
      log("Web Completed!");
    } else if (Platform.isAndroid || Platform.isIOS) {
      await DocumentFileSavePlus.saveFile(uint8, downloadFileName, mimeType);
      if (Platform.isAndroid) {
        if (mounted) {
          const String msg = "The file saved successfully at Downloads folder.";
          const SnackBar snack = SnackBar(content: Text(msg));
          ScaffoldMessenger.of(context).showSnackBar(snack);
          log("Android Completed!");
        } else {}
      } else if (Platform.isIOS) {
        log("iOS Completed!");
      } else {}
    } else {
      log("Unsupported Platform");
    }
    return Future<void>.value();
  }

  CellIndex indexByColumnRow({required int cIndex, required int rIndex}) {
    return CellIndex.indexByColumnRow(columnIndex: cIndex, rowIndex: rIndex);
  }
}
