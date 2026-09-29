import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/presentation/widgets/draft_preview_canvas.dart';

void main() {
  group('DraftPreviewCanvas Widget Tests', () {
    testWidgets('renders header title and view mode action icons', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: DraftPreviewCanvas(),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Draft Preview Canvas'), findsOneWidget);
      expect(find.byIcon(Icons.calendar_view_month), findsOneWidget);
      expect(find.byIcon(Icons.table_chart_outlined), findsOneWidget);
    });

    testWidgets('renders "Awaiting Generation" empty state when isGenerated is false', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: DraftPreviewCanvas(
              isGenerated: false,
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Awaiting Generation'), findsOneWidget);
      expect(
        find.textContaining('Click "Generate Schedule" above or import an Excel spreadsheet'),
        findsOneWidget,
      );
      expect(find.byIcon(Icons.calendar_month_outlined), findsOneWidget);
    });

    testWidgets('renders generated draft content when isGenerated is true', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: DraftPreviewCanvas(
              isGenerated: true,
              generatedContent: Text('Draft Schedule Table Placeholder'),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Draft Preview Canvas'), findsOneWidget);
      expect(find.text('Awaiting Generation'), findsNothing);
      expect(find.text('Draft Schedule Table Placeholder'), findsOneWidget);
    });
  });
}
