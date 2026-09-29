import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/presentation/widgets/generation_mode_dialog.dart';

void main() {
  group('GenerationModeDialog Widget Tests', () {
    testWidgets('renders dialog title, both generation options, and cancel button', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: GenerationModeDialog(),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Select Generation Mode'), findsOneWidget);
      expect(find.text('Current Department Roster'), findsOneWidget);
      expect(find.text('Import from Excel (.xlsx)'), findsOneWidget);
      expect(find.text('Cancel'), findsOneWidget);

      expect(find.byIcon(Icons.people_alt_outlined), findsOneWidget);
      expect(find.byIcon(Icons.upload_file_outlined), findsOneWidget);
    });

    testWidgets('tapping Current Department Roster pops dialog with GenerationMode.currentRoster', (tester) async {
      GenerationMode? selectedMode;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: Builder(
              builder: (context) => ElevatedButton(
                onPressed: () async {
                  selectedMode = await showDialog<GenerationMode>(
                    context: context,
                    builder: (_) => const GenerationModeDialog(),
                  );
                },
                child: const Text('Open Dialog'),
              ),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // Open dialog
      await tester.tap(find.text('Open Dialog'));
      await tester.pumpAndSettle();

      // Tap current roster option
      await tester.tap(find.text('Current Department Roster'));
      await tester.pumpAndSettle();

      expect(selectedMode, equals(GenerationMode.currentRoster));
      expect(find.byType(GenerationModeDialog), findsNothing);
    });

    testWidgets('tapping Import from Excel pops dialog with GenerationMode.excelUpload', (tester) async {
      GenerationMode? selectedMode;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: Builder(
              builder: (context) => ElevatedButton(
                onPressed: () async {
                  selectedMode = await showDialog<GenerationMode>(
                    context: context,
                    builder: (_) => const GenerationModeDialog(),
                  );
                },
                child: const Text('Open Dialog'),
              ),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // Open dialog
      await tester.tap(find.text('Open Dialog'));
      await tester.pumpAndSettle();

      // Tap excel option
      await tester.tap(find.text('Import from Excel (.xlsx)'));
      await tester.pumpAndSettle();

      expect(selectedMode, equals(GenerationMode.excelUpload));
      expect(find.byType(GenerationModeDialog), findsNothing);
    });

    testWidgets('tapping Cancel dismisses dialog returning null', (tester) async {
      GenerationMode? selectedMode = GenerationMode.currentRoster;

      await tester.pumpWidget(
        MaterialApp(
          home: Scaffold(
            body: Builder(
              builder: (context) => ElevatedButton(
                onPressed: () async {
                  selectedMode = await showDialog<GenerationMode>(
                    context: context,
                    builder: (_) => const GenerationModeDialog(),
                  );
                },
                child: const Text('Open Dialog'),
              ),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // Open dialog
      await tester.tap(find.text('Open Dialog'));
      await tester.pumpAndSettle();

      // Tap Cancel
      await tester.tap(find.text('Cancel'));
      await tester.pumpAndSettle();

      expect(selectedMode, isNull);
      expect(find.byType(GenerationModeDialog), findsNothing);
    });
  });
}
