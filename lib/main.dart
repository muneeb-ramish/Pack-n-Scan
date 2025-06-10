
// ignore_for_file: library_private_types_in_public_api, unused_local_variable

import 'package:flutter/material.dart';
import 'package:excel/excel.dart' as excel;
import 'package:file_picker/file_picker.dart';
import 'package:intl/intl.dart';
import 'package:window_manager/window_manager.dart';
import 'dart:io';
import 'dart:async';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await windowManager.ensureInitialized();
  WindowOptions windowOptions = const WindowOptions(
    size: Size(800, 600),
    center: true,
    backgroundColor: Colors.transparent,
    skipTaskbar: false,
    titleBarStyle: TitleBarStyle.normal,
  );
  windowManager.waitUntilReadyToShow(windowOptions, () async {
    await windowManager.maximize();
    await windowManager.show();
    await windowManager.focus();
  });

  runApp(const RFIDLoggerApp());
}

class RFIDLoggerApp extends StatelessWidget {
  const RFIDLoggerApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'RFID Master Data Logging',
      theme: ThemeData(
        primaryColor: const Color(0xFF0078D4),
        scaffoldBackgroundColor: const Color(0xFFF0F8FF),
        fontFamily: 'Segoe UI',
      ),
      debugShowCheckedModeBanner: false,
      home: const RFIDLoggerHome(),
    );
  }
}

class RFIDLoggerHome extends StatefulWidget {
  const RFIDLoggerHome({super.key});

  @override
  _RFIDLoggerHomeState createState() => _RFIDLoggerHomeState();
}

class _RFIDLoggerHomeState extends State<RFIDLoggerHome> {
  final TextEditingController _inputController = TextEditingController();
  final TextEditingController _orderNumberController = TextEditingController();
  final TextEditingController _customerController = TextEditingController();
  final TextEditingController _barcodeLabelController = TextEditingController();
  final TextEditingController _articleController = TextEditingController();
  final TextEditingController _totalPOQtyController = TextEditingController();
  final TextEditingController _boxTypeController = TextEditingController();
  final TextEditingController _boxCapacityController = TextEditingController();
  final TextEditingController _wmsIdController = TextEditingController();

  final Map<String, int> _uidDict = {};
  final List<String> _logEntries = [];
  final ValueNotifier<int> _uniqueCountNotifier = ValueNotifier<int>(0);
  int _rowCount = 2; // Start row count from 2 for EPC UID in column I
  String _statusMessage = 'Ready to scan...';
  late excel.Excel _excel;
  late FocusNode _textFieldFocusNode;
  final List<String> _pendingTags = [];
  bool _isProcessing = false;

  @override
  void initState() {
    super.initState();
    _initializeExcel();
    _textFieldFocusNode = FocusNode();
    WidgetsBinding.instance.addPostFrameCallback((_) {
      _textFieldFocusNode.requestFocus();
    });
  }

  @override
  void dispose() {
    _textFieldFocusNode.dispose();
    _inputController.dispose();
    _orderNumberController.dispose();
    _customerController.dispose();
    _barcodeLabelController.dispose();
    _articleController.dispose();
    _totalPOQtyController.dispose();
    _boxTypeController.dispose();
    _boxCapacityController.dispose();
    _wmsIdController.dispose();
    _uniqueCountNotifier.dispose();
    super.dispose();
  }

  void _initializeExcel() {
    _excel = excel.Excel.createExcel();
    excel.Sheet sheet = _excel['Sheet1'];

    // Titles in Row 1
    sheet.cell(excel.CellIndex.indexByString('A1')).value = excel.TextCellValue('Customer');
    sheet.cell(excel.CellIndex.indexByString('B1')).value = excel.TextCellValue('Order No');
    sheet.cell(excel.CellIndex.indexByString('C1')).value = excel.TextCellValue('Article');
    sheet.cell(excel.CellIndex.indexByString('D1')).value = excel.TextCellValue('Total PO Qty');
    sheet.cell(excel.CellIndex.indexByString('E1')).value = excel.TextCellValue('Barcode Label');
    sheet.cell(excel.CellIndex.indexByString('F1')).value = excel.TextCellValue('Box Type');
    sheet.cell(excel.CellIndex.indexByString('G1')).value = excel.TextCellValue('Box Capacity');
    sheet.cell(excel.CellIndex.indexByString('H1')).value = excel.TextCellValue('WMS ID');
    sheet.cell(excel.CellIndex.indexByString('I1')).value = excel.TextCellValue('EPC UID');

    // Initial values in Row 2
    _updateExcelMetadata();
  }

  void _updateExcelMetadata() {
    excel.Sheet sheet = _excel['Sheet1'];
    sheet.cell(excel.CellIndex.indexByString('A2')).value = excel.TextCellValue(
        _customerController.text.trim().isEmpty ? 'NoCustomer' : _customerController.text.trim());
    sheet.cell(excel.CellIndex.indexByString('B2')).value = excel.TextCellValue(
        _orderNumberController.text.trim().isEmpty ? 'NoOrder' : _orderNumberController.text.trim());
    sheet.cell(excel.CellIndex.indexByString('C2')).value = excel.TextCellValue(
        _articleController.text.trim().isEmpty ? 'NoArticle' : _articleController.text.trim());
    sheet.cell(excel.CellIndex.indexByString('D2')).value = excel.TextCellValue(
        _totalPOQtyController.text.trim().isEmpty ? 'NoQty' : _totalPOQtyController.text.trim());
    sheet.cell(excel.CellIndex.indexByString('E2')).value = excel.TextCellValue(
        _barcodeLabelController.text.trim().isEmpty ? 'NoLabel' : _barcodeLabelController.text.trim());
    sheet.cell(excel.CellIndex.indexByString('F2')).value = excel.TextCellValue(
        _boxTypeController.text.trim().isEmpty ? 'NoType' : _boxTypeController.text.trim());
    sheet.cell(excel.CellIndex.indexByString('G2')).value = excel.TextCellValue(
        _boxCapacityController.text.trim().isEmpty ? 'NoCapacity' : _boxCapacityController.text.trim());
    sheet.cell(excel.CellIndex.indexByString('H2')).value = excel.TextCellValue(
        _wmsIdController.text.trim().isEmpty ? 'NoWMSID' : _wmsIdController.text.trim());
  }

  void _processTag(String value) async {
    if (_isProcessing) {
      _pendingTags.add(value);
      return;
    }

    _isProcessing = true;
    String uid = value.trim();

    // Skip values starting with "http" or "https"
    if (RegExp(r'^https?://', caseSensitive: false).hasMatch(uid)) {
      _statusMessage = 'Skipped URL starting with http or https (any case)';
      _inputController.clear();
      _textFieldFocusNode.requestFocus();
      setState(() {});
      _isProcessing = false;
      if (_pendingTags.isNotEmpty) {
        String nextTag = _pendingTags.removeAt(0);
        _processTag(nextTag);
      }
      return;
    }

    if (uid.length >= 15) {
      if (!_uidDict.containsKey(uid)) {
        int serialNo = _uidDict.length + 1; // Serial number starts at 1
        String timestamp = DateFormat('yyyy-MM-dd HH:mm:ss').format(DateTime.now());
        String logLine = '$serialNo\t$uid\t$timestamp'; // Include Serial No and Timestamp in log

        _uidDict[uid] = _rowCount - 1; // Index for uniqueness tracking
        _logEntries.insert(0, logLine);
        _excel['Sheet1'].cell(excel.CellIndex.indexByString('I$_rowCount')).value = excel.TextCellValue(uid);

        _rowCount++;
        _uniqueCountNotifier.value = _uidDict.length;
        _statusMessage = 'Logged new UID. Total unique UIDs: ${_uidDict.length}';
      }
    }

    _inputController.clear();
    _textFieldFocusNode.requestFocus();
    setState(() {});

    _isProcessing = false;
    if (_pendingTags.isNotEmpty) {
      String nextTag = _pendingTags.removeAt(0);
      _processTag(nextTag);
    }
  }

  Future<void> _saveFile() async {
    String? folderPath = await FilePicker.platform.getDirectoryPath();
    if (folderPath != null) {
      String orderNumber = _orderNumberController.text.trim().isEmpty ? 'NoOrder' : _orderNumberController.text.trim();
      String customer = _customerController.text.trim().isEmpty ? 'NoCustomer' : _customerController.text.trim();
      String timestamp = DateFormat('yyyy-MM-dd-HH-mm-ss').format(DateTime.now());
      String fileName = '${orderNumber}_${customer}_RFID_Log_$timestamp.xlsx';
      String filePath = '$folderPath\\$fileName';

      // Update metadata in Excel before saving
      _updateExcelMetadata();

      try {
        File(filePath)
          ..createSync(recursive: true)
          ..writeAsBytesSync(_excel.encode()!);
        setState(() {
          _statusMessage = 'File saved successfully to: $filePath';
        });
      } catch (e) {
        setState(() {
          _statusMessage = 'Error saving file: $e';
        });
      }
    } else {
      setState(() {
        _statusMessage = 'Save cancelled by user';
      });
    }
    _textFieldFocusNode.requestFocus();
  }

  void _clearList() {
    setState(() {
      _uidDict.clear();
      _logEntries.clear();
      _rowCount = 2; // Reset row count to start at I2
      _excel = excel.Excel.createExcel();
      excel.Sheet sheet = _excel['Sheet1'];

      // Clear all text box values
      _inputController.clear();
      _orderNumberController.clear();
      _customerController.clear();
      _barcodeLabelController.clear();
      _articleController.clear();
      _totalPOQtyController.clear();
      _boxTypeController.clear();
      _boxCapacityController.clear();
      _wmsIdController.clear();

      // Reinitialize Excel with titles and cleared metadata
      sheet.cell(excel.CellIndex.indexByString('A1')).value = excel.TextCellValue('Customer');
      sheet.cell(excel.CellIndex.indexByString('B1')).value = excel.TextCellValue('Order No');
      sheet.cell(excel.CellIndex.indexByString('C1')).value = excel.TextCellValue('Article');
      sheet.cell(excel.CellIndex.indexByString('D1')).value = excel.TextCellValue('Total PO Qty');
      sheet.cell(excel.CellIndex.indexByString('E1')).value = excel.TextCellValue('Barcode Label');
      sheet.cell(excel.CellIndex.indexByString('F1')).value = excel.TextCellValue('Box Type');
      sheet.cell(excel.CellIndex.indexByString('G1')).value = excel.TextCellValue('Box Capacity');
      sheet.cell(excel.CellIndex.indexByString('H1')).value = excel.TextCellValue('WMS ID');
      sheet.cell(excel.CellIndex.indexByString('I1')).value = excel.TextCellValue('EPC UID');

      _updateExcelMetadata();

      _uniqueCountNotifier.value = 0;
      _statusMessage = 'List cleared. Ready for new acquisitions.';
    });
    _textFieldFocusNode.requestFocus();
  }

  void _exitApp() {
    exit(0);
  }

  @override
  Widget build(BuildContext context) {
    final screenWidth = MediaQuery.of(context).size.width;
    final maxContentWidth = screenWidth - 40;

    return Scaffold(
      body: Padding(
        padding: const EdgeInsets.all(20.0),
        child: Column(
          children: [
            Row(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                const Text(
                  ' RFID Master Data Logging',
                  style: TextStyle(
                    fontSize: 24,
                    color: Color(0xFF0078D4),
                    fontWeight: FontWeight.bold,
                  ),
                ),
              ],
            ),
            const SizedBox(height: 10),
            Row(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                const Text(
                  'Project: Pack\'n\'Scan',
                  style: TextStyle(fontSize: 18, color: Color(0xFF0078D4)),
                ),
              ],
            ),
            const SizedBox(height: 10),
            Row(
              crossAxisAlignment: CrossAxisAlignment.start,
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                Padding(padding: EdgeInsets.only(left: 100)),
                SizedBox(
                  width: 300,
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      const Text('Customer', style: TextStyle(fontSize: 16, color: Color(0xFF0078D4))),
                      const SizedBox(height: 5),
                      TextField(
                        controller: _customerController,
                        decoration: InputDecoration(
                          hintText: 'Enter Customer',
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                          focusedBorder: OutlineInputBorder(
                            borderRadius: BorderRadius.circular(8),
                            borderSide: const BorderSide(color: Color(0xFF005A8B)),
                          ),
                        ),
                      ),
                      const SizedBox(height: 10),
                      const Text('Order Number', style: TextStyle(fontSize: 16, color: Color(0xFF0078D4))),
                      const SizedBox(height: 5),
                      TextField(
                        controller: _orderNumberController,
                        decoration: InputDecoration(
                          hintText: 'Enter Order Number',
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                          focusedBorder: OutlineInputBorder(
                            borderRadius: BorderRadius.circular(8),
                            borderSide: const BorderSide(color: Color(0xFF005A8B)),
                          ),
                        ),
                      ),
                      const SizedBox(height: 10),
                      const Text('Article', style: TextStyle(fontSize: 16, color: Color(0xFF0078D4))),
                      const SizedBox(height: 5),
                      TextField(
                        controller: _articleController,
                        decoration: InputDecoration(
                          hintText: 'Enter Article',
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                          focusedBorder: OutlineInputBorder(
                            borderRadius: BorderRadius.circular(8),
                            borderSide: const BorderSide(color: Color(0xFF005A8B)),
                          ),
                        ),
                      ),
                      const SizedBox(height: 10),
                      const Text('Total PO Qty', style: TextStyle(fontSize: 16, color: Color(0xFF0078D4))),
                      const SizedBox(height: 5),
                      TextField(
                        controller: _totalPOQtyController,
                        decoration: InputDecoration(
                          hintText: 'Enter Total PO Qty',
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                          focusedBorder: OutlineInputBorder(
                            borderRadius: BorderRadius.circular(8),
                            borderSide: const BorderSide(color: Color(0xFF005A8B)),
                          ),
                        ),
                      ),
                      const SizedBox(height: 10),
                      const Text('Barcode Label', style: TextStyle(fontSize: 16, color: Color(0xFF0078D4))),
                      const SizedBox(height: 5),
                      TextField(
                        controller: _barcodeLabelController,
                        decoration: InputDecoration(
                          hintText: 'Enter Barcode Label',
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                          focusedBorder: OutlineInputBorder(
                            borderRadius: BorderRadius.circular(8),
                            borderSide: const BorderSide(color: Color(0xFF005A8B)),
                          ),
                        ),
                      ),
                      const SizedBox(height: 10),
                      const Text('Box Type', style: TextStyle(fontSize: 16, color: Color(0xFF0078D4))),
                      const SizedBox(height: 5),
                      TextField(
                        controller: _boxTypeController,
                        decoration: InputDecoration(
                          hintText: 'Enter Box Type',
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                          focusedBorder: OutlineInputBorder(
                            borderRadius: BorderRadius.circular(8),
                            borderSide: const BorderSide(color: Color(0xFF005A8B)),
                          ),
                        ),
                      ),
                      const SizedBox(height: 10),
                      const Text('Box Capacity', style: TextStyle(fontSize: 16, color: Color(0xFF0078D4))),
                      const SizedBox(height: 5),
                      TextField(
                        controller: _boxCapacityController,
                        decoration: InputDecoration(
                          hintText: 'Enter Box Capacity',
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                          focusedBorder: OutlineInputBorder(
                            borderRadius: BorderRadius.circular(8),
                            borderSide: const BorderSide(color: Color(0xFF005A8B)),
                          ),
                        ),
                      ),
                      const SizedBox(height: 10),
                      const Text('WMS ID', style: TextStyle(fontSize: 16, color: Color(0xFF0078D4))),
                      const SizedBox(height: 5),
                      TextField(
                        controller: _wmsIdController,
                        decoration: InputDecoration(
                          hintText: 'Enter WMS ID',
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                          focusedBorder: OutlineInputBorder(
                            borderRadius: BorderRadius.circular(8),
                            borderSide: const BorderSide(color: Color(0xFF005A8B)),
                          ),
                        ),
                      ),
                    ],
                  ),
                ),
                const SizedBox(width: 20),
                Expanded(
                  child: Padding(
                    padding: const EdgeInsets.only(left: 25),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        SizedBox(
                          width: maxContentWidth > 600 ? 600 : maxContentWidth - 200,
                          child: TextField(
                            controller: _inputController,
                            focusNode: _textFieldFocusNode,
                            decoration: InputDecoration(
                              hintText: 'Scan or Enter UID',
                              border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                              focusedBorder: OutlineInputBorder(
                                borderRadius: BorderRadius.circular(8),
                                borderSide: const BorderSide(color: Color(0xFF005A8B)),
                              ),
                            ),
                            onSubmitted: (value) {
                              _processTag(value);
                            },
                            autofocus: true,
                          ),
                        ),
                        const SizedBox(height: 20),
                        Container(
                          width: maxContentWidth > 600 ? 600 : maxContentWidth - 320,
                          height: 400,
                          decoration: BoxDecoration(
                            border: Border.all(color: const Color(0xFF0078D4)),
                            borderRadius: BorderRadius.circular(8),
                            color: Colors.white,
                          ),
                          child: SingleChildScrollView(
                            child: Padding(
                              padding: const EdgeInsets.all(10.0),
                              child: Text(
                                _logEntries.join('\n'),
                                style: const TextStyle(fontFamily: 'Courier New'),
                              ),
                            ),
                          ),
                        ),
                        const SizedBox(height: 20),
                        ValueListenableBuilder<int>(
                          valueListenable: _uniqueCountNotifier,
                          builder: (context, count, child) {
                            return Text(_statusMessage, style: const TextStyle(fontSize: 16));
                          },
                        ),
                        const SizedBox(height: 20),
                        Padding(
                          padding: const EdgeInsets.fromLTRB(27, 0, 0, 0),
                          child: Row(
                            children: [
                              ElevatedButton(
                                onPressed: _saveFile,
                                style: ElevatedButton.styleFrom(
                                  backgroundColor: const Color(0xFF0078D4),
                                  padding: const EdgeInsets.symmetric(horizontal: 40, vertical: 18),
                                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                                  minimumSize: const Size(180, 60),
                                ),
                                child: const Text('Save', style: TextStyle(color: Colors.white, fontSize: 18)),
                              ),
                              const SizedBox(width: 15),
                              ElevatedButton(
                                onPressed: _clearList,
                                style: ElevatedButton.styleFrom(
                                  backgroundColor: const Color(0xFF0078D4),
                                  padding: const EdgeInsets.symmetric(horizontal: 40, vertical: 18),
                                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                                  minimumSize: const Size(180, 60),
                                ),
                                child: const Text('Clear List', style: TextStyle(color: Colors.white, fontSize: 18)),
                              ),
                              const SizedBox(width: 15),
                              ElevatedButton(
                                onPressed: _exitApp,
                                style: ElevatedButton.styleFrom(
                                  backgroundColor: const Color(0xFF0078D4),
                                  padding: const EdgeInsets.symmetric(horizontal: 40, vertical: 18),
                                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                                  minimumSize: const Size(180, 60),
                                ),
                                child: const Text('Exit', style: TextStyle(color: Colors.white, fontSize: 18)),
                              ),
                            ],
                          ),
                        ),
                      ],
                    ),
                  ),
                ),
              ],
            ),
            const Spacer(),
            const Text('© INTERLOOP LIMITED', style: TextStyle(fontSize: 14, color: Color(0xFF888888))),
          ],
        ),
      ),
    );
  }
}