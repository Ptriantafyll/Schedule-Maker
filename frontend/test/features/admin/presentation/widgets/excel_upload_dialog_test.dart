import 'dart:typed_data';
// ignore: depend_on_referenced_packages
import 'package:cross_file/cross_file.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/presentation/widgets/excel_upload_dialog.dart';

base class FakePlatformFile extends PlatformFile {
  FakePlatformFile({
    required this.name,
    required this.fileSize,
  });

  @override
  final String name;

  final int fileSize;

  @override
  Uri get uri => Uri.file(name);

  @override
  XFile get xFile => XFile(name);

  @override
  int? lengthSync() => fileSize;

  @override
  Future<int?> length() async => fileSize;

  @override
  Future<Uint8List> readAsBytes() async => Uint8List(0);

  @override
  Stream<Uint8List> readAsByteStream() => const Stream.empty();
}

void main() {
  final testFile = FakePlatformFile(
    name: 'november_roster.xlsx',
    fileSize: 25600, // 25.0 KB
  );

  group('ExcelUploadDialog Widget Tests', () {
    testWidgets('renders title, upload prompt, cancel button, and disabled generate button', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: ExcelUploadDialog(),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Import Schedule Workbook'), findsOneWidget);
      expect(find.text('Click to select .xlsx file'), findsOneWidget);
      expect(find.text('Cancel'), findsOneWidget);

      final generateButtonFinder = find.widgetWithText(FilledButton, 'Generate Schedule');
      expect(generateButtonFinder, findsOneWidget);

      final generateButton = tester.widget<FilledButton>(generateButtonFinder);
      expect(generateButton.onPressed, isNull);
    });

    testWidgets('picking a file displays filename, formatted size, and enables generate button', (tester) async {
      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: ExcelUploadDialog(
              onPickFile: () async => testFile,
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // Tap the browse upload container
      await tester.tap(find.text('Click to select .xlsx file'));
      await tester.pumpAndSettle();

      // Selected file details should now be visible
      expect(find.text('november_roster.xlsx'), findsOneWidget);
      expect(find.textContaining('25.0 KB'), findsOneWidget);
      expect(find.byIcon(Icons.close), findsOneWidget);

      // Generate button should now be enabled
      final generateButton = tester.widget<FilledButton>(
        find.widgetWithText(FilledButton, 'Generate Schedule'),
      );
      expect(generateButton.onPressed, isNotNull);
    });

    testWidgets('tapping remove icon clears selected file and disables generate button again', (tester) async {
      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: ExcelUploadDialog(
              onPickFile: () async => testFile,
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // Pick file
      await tester.tap(find.text('Click to select .xlsx file'));
      await tester.pumpAndSettle();
      expect(find.text('november_roster.xlsx'), findsOneWidget);

      // Tap remove icon
      await tester.tap(find.byIcon(Icons.close));
      await tester.pumpAndSettle();

      // Should return to prompt and disabled state
      expect(find.text('november_roster.xlsx'), findsNothing);
      expect(find.text('Click to select .xlsx file'), findsOneWidget);

      final generateButton = tester.widget<FilledButton>(
        find.widgetWithText(FilledButton, 'Generate Schedule'),
      );
      expect(generateButton.onPressed, isNull);
    });

    testWidgets('tapping Generate Schedule pops dialog returning the picked PlatformFile', (tester) async {
      PlatformFile? returnedFile;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: Builder(
              builder: (context) => ElevatedButton(
                onPressed: () async {
                  returnedFile = await showDialog<PlatformFile>(
                    context: context,
                    builder: (_) => ExcelUploadDialog(
                      onPickFile: () async => testFile,
                    ),
                  );
                },
                child: const Text('Open Modal'),
              ),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // Open dialog
      await tester.tap(find.text('Open Modal'));
      await tester.pumpAndSettle();

      // Pick file
      await tester.tap(find.text('Click to select .xlsx file'));
      await tester.pumpAndSettle();

      // Tap Generate Schedule
      await tester.tap(find.widgetWithText(FilledButton, 'Generate Schedule'));
      await tester.pumpAndSettle();

      expect(returnedFile, equals(testFile));
      expect(find.byType(ExcelUploadDialog), findsNothing);
    });

    testWidgets('tapping Cancel dismisses dialog returning null', (tester) async {
      PlatformFile? returnedFile = testFile;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: Builder(
              builder: (context) => ElevatedButton(
                onPressed: () async {
                  returnedFile = await showDialog<PlatformFile>(
                    context: context,
                    builder: (_) => const ExcelUploadDialog(),
                  );
                },
                child: const Text('Open Modal'),
              ),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // Open dialog
      await tester.tap(find.text('Open Modal'));
      await tester.pumpAndSettle();

      // Tap Cancel
      await tester.tap(find.text('Cancel'));
      await tester.pumpAndSettle();

      expect(returnedFile, isNull);
      expect(find.byType(ExcelUploadDialog), findsNothing);
    });
  });
}
