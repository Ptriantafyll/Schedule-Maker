import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/presentation/widgets/admin_metric_cards.dart';

void main() {
  group('AdminMetricCards Widget Tests (4-Card Pre-Flight System)', () {
    testWidgets('renders Schedule Due Date card with due date, countdown, and submission status', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: AdminMetricCards(
              dueDate: 'Oct 25',
              daysLeft: 5,
              submissionStatus: 'Requests Closed',
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Schedule Due Date'), findsOneWidget);
      expect(find.text('Oct 25'), findsOneWidget);
      expect(find.text('5 Days Left'), findsOneWidget);
      expect(find.text('Requests Closed'), findsOneWidget);
    });

    testWidgets('renders Pending Approvals card with count and action status', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: AdminMetricCards(
              pendingApprovalsCount: 2,
              pendingApprovalsStatus: 'Action Required',
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Pending Approvals'), findsOneWidget);
      expect(find.text('2'), findsOneWidget);
      expect(find.text('Action Required'), findsOneWidget);
    });

    testWidgets('renders Staff Pool & Capacity card with doctor count, total shifts, and ratio', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: AdminMetricCards(
              availableStaffCount: 18,
              totalShiftsCount: 90,
              staffCapacitySubtitle: '~5.0 shifts/doctor',
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Staff Pool & Capacity'), findsOneWidget);
      expect(find.text('18 Doctors'), findsOneWidget);
      expect(find.text('90 Shifts to Assign'), findsOneWidget);
      expect(find.text('~5.0 shifts/doctor'), findsOneWidget);
    });

    testWidgets('renders Coverage Bottlenecks card with tight days count, warning badge, and chips', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: AdminMetricCards(
              bottleneckDaysCount: 3,
              bottleneckStatus: 'Requires Attention',
              bottleneckDates: ['ICU - Nov 12', 'ER - Nov 15'],
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      expect(find.text('Coverage Bottlenecks'), findsOneWidget);
      expect(find.text('3 Tight Days'), findsOneWidget);
      expect(find.text('Requires Attention'), findsOneWidget);
      expect(find.text('ICU - Nov 12'), findsOneWidget);
      expect(find.text('ER - Nov 15'), findsOneWidget);
    });

    testWidgets('renders all 4 cards with default constructor values', (tester) async {
      await tester.pumpWidget(
        const MaterialApp(
          home: Scaffold(
            body: SingleChildScrollView(
              child: AdminMetricCards(),
            ),
          ),
        ),
      );
      await tester.pumpAndSettle();

      // All 4 card titles present
      expect(find.text('Schedule Due Date'), findsOneWidget);
      expect(find.text('Pending Approvals'), findsOneWidget);
      expect(find.text('Staff Pool & Capacity'), findsOneWidget);
      expect(find.text('Coverage Bottlenecks'), findsOneWidget);

      // Default metric indicators present
      expect(find.text('Oct 25'), findsOneWidget);
      expect(find.text('2'), findsOneWidget);
      expect(find.text('18 Doctors'), findsOneWidget);
      expect(find.text('3 Tight Days'), findsOneWidget);
    });
  });
}
