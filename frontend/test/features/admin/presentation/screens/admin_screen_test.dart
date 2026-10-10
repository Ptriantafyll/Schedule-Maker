import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/data/repositories/schedule_draft_repository.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';
import 'package:frontend/features/admin/presentation/screens/admin_screen.dart';
import 'package:frontend/features/auth/domain/models/user.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';
import 'package:frontend/features/auth/presentation/controllers/auth_controller.dart';
import 'package:frontend/features/admin/presentation/widgets/excel_upload_dialog.dart';
import 'package:frontend/features/admin/presentation/widgets/generation_mode_dialog.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/domain/models/schedule_summary.dart';
import 'package:frontend/features/admin/domain/models/target_month_info.dart';
import 'package:frontend/shared/widgets/bottom_nav_bar.dart';
import 'package:frontend/shared/widgets/profile_drawer.dart';

class FakeAdminScheduleDraftRepository implements ScheduleDraftRepository {
  ScheduleDraft? draftToReturn;
  List<int> exportBytesToReturn = [1, 2, 3];
  Object? errorToThrow;
  int getActiveDraftCallCount = 0;
  int exportExcelCallCount = 0;
  int generateFromExcelCallCount = 0;
  int generateFromRosterCallCount = 0;
  int publishScheduleDraftCallCount = 0;
  List<ScheduleSummary> historyToReturn = const [];

  @override
  Future<ScheduleDraft> generateFromExcel({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
  }) async {
    generateFromExcelCallCount++;
    if (errorToThrow != null) throw errorToThrow!;
    return draftToReturn ?? _createSampleDraft();
  }

  @override
  Future<ScheduleDraft?> getActiveDraft({
    required String targetMonth,
    String? departmentId,
  }) async {
    getActiveDraftCallCount++;
    if (errorToThrow != null) throw errorToThrow!;
    return draftToReturn;
  }

  String? capturedDraftId;

  @override
  Future<List<int>> exportExcel({
    required String draftId,
  }) async {
    exportExcelCallCount++;
    capturedDraftId = draftId;
    return exportBytesToReturn;
  }

  @override
  Future<void> generateFromRoster({required String month}) async {
    generateFromRosterCallCount++;
  }

  @override
  Future<List<ScheduleSummary>> fetchScheduleHistory({String? departmentId}) async {
    return historyToReturn;
  }

  @override
  Future<TargetMonthInfo> fetchTargetMonthInfo({String? departmentId}) async {
    return const TargetMonthInfo(nextTargetMonth: '2026-11');
  }

  @override
  Future<ScheduleDraft> publishScheduleDraft({required String draftId}) async {
    publishScheduleDraftCallCount++;
    if (errorToThrow != null) throw errorToThrow!;
    return draftToReturn ?? _createSampleDraft();
  }

  static ScheduleDraft _createSampleDraft() {
    return ScheduleDraft(
      id: 'draft-101',
      departmentId: 'dept-er',
      targetMonth: '2026-11',
      sourceFilename: 'november_roster.xlsx',
      totalDuties: 28,
      solverStatus: 'OPTIMAL',
      status: 'draft',
      assignments: const [
        ScheduleAssignment(
          date: '2026-11-01',
          dayName: 'Sunday',
          doctorName: 'Dr. Gregory House',
          doctorEmail: 'house@hospital.org',
          position: 'ER',
          shift: 'Night',
        ),
      ],
      unavailabilities: const {},
      createdAt: DateTime.parse('2026-11-01T08:00:00.000Z'),
      updatedAt: DateTime.parse('2026-11-01T08:00:00.000Z'),
    );
  }
}

void main() {
  const testAdmin = User(
    id: 'adm-101',
    email: 'admin@hospital.org',
    fullName: 'Dr. Gregory House',
    role: UserRole.departmentAdmin,
    departmentId: 'dept-er',
  );

  Widget createWidgetUnderTest({
    User? user = testAdmin,
    ScheduleDraftRepository? repository,
    Future<String?> Function({
      required String fileName,
      required List<int> bytes,
    })? onSaveFile,
  }) {
    return ProviderScope(
      overrides: [
        currentUserProvider.overrideWithValue(user),
        scheduleDraftRepositoryProvider.overrideWithValue(
          repository ?? FakeAdminScheduleDraftRepository(),
        ),
      ],
      child: MaterialApp(
        home: AdminScreen(
          onSaveFile: onSaveFile ??
              ({required fileName, required bytes}) async =>
                  '/downloads/$fileName',
        ),
      ),
    );
  }

  group('AdminScreen Scaffold & Shell Tests', () {
    testWidgets('renders AppBar with branding, profile avatar, and notification icon', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      expect(find.text('MedShift Admin'), findsOneWidget);
      expect(
        find.descendant(of: find.byType(AppBar), matching: find.byType(CircleAvatar)),
        findsOneWidget,
      );
      expect(
        find.byWidgetPredicate(
          (widget) => widget is Icon && (widget.icon == Icons.notifications_none || widget.icon == Icons.notifications_outlined || widget.icon == Icons.notifications),
        ),
        findsOneWidget,
      );
      expect(find.byType(BottomNavBar), findsOneWidget);
    });

    testWidgets('tapping circular profile avatar opens drawer showing admin full name', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      // Initially drawer is closed
      expect(find.byType(ProfileDrawer), findsNothing);

      // Tap profile avatar in AppBar
      final avatarFinder = find.descendant(of: find.byType(AppBar), matching: find.byType(CircleAvatar));
      expect(avatarFinder, findsOneWidget);
      await tester.tap(avatarFinder);
      await tester.pumpAndSettle();

      // Drawer should now be open
      expect(find.byType(ProfileDrawer), findsOneWidget);
      expect(find.text('Dr. Gregory House'), findsOneWidget);
    });

    testWidgets('tapping bottom navigation tabs updates selected tab', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      // Tap Schedule tab
      await tester.tap(find.text('Schedule'));
      await tester.pumpAndSettle();

      // Tap Dashboard tab
      await tester.tap(find.text('Dashboard'));
      await tester.pumpAndSettle();

      // Tap Admin tab
      await tester.tap(find.text('Admin'));
      await tester.pumpAndSettle();
    });

    testWidgets('tapping Generate Schedule button opens GenerationModeDialog', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      expect(find.byType(GenerationModeDialog), findsNothing);

      await tester.tap(find.text('Generate Schedule'));
      await tester.pumpAndSettle();

      expect(find.byType(GenerationModeDialog), findsOneWidget);
      expect(find.text('Select Generation Mode'), findsOneWidget);
    });

    testWidgets('selecting Current Department Roster generates schedule and updates canvas and snackbar', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      // Initially in awaiting generation state
      expect(find.text('Awaiting Generation'), findsOneWidget);
      expect(find.text('Draft Schedule Generated'), findsNothing);

      // Tap Generate Schedule on Hero Card
      await tester.tap(find.text('Generate Schedule'));
      await tester.pumpAndSettle();

      // Select Current Department Roster
      await tester.tap(find.text('Current Department Roster'));
      await tester.pumpAndSettle();

      // Draft canvas should now display Draft Schedule Generated
      expect(find.text('Draft Schedule Generated'), findsOneWidget);
      expect(find.text('Awaiting Generation'), findsNothing);

      // SnackBar should be displayed
      expect(find.text('Schedule generated from current roster'), findsOneWidget);
    });

    testWidgets('selecting Import from Excel opens ExcelUploadDialog', (tester) async {
      await tester.pumpWidget(createWidgetUnderTest());
      await tester.pumpAndSettle();

      expect(find.byType(ExcelUploadDialog), findsNothing);

      // Tap Generate Schedule on Hero Card
      await tester.tap(find.text('Generate Schedule'));
      await tester.pumpAndSettle();

      // Select Import from Excel (.xlsx)
      await tester.tap(find.text('Import from Excel (.xlsx)'));
      await tester.pumpAndSettle();

      // ExcelUploadDialog should be displayed
      expect(find.byType(ExcelUploadDialog), findsOneWidget);
      expect(find.text('Import Schedule Workbook'), findsOneWidget);

      // Cancel Excel dialog
      await tester.tap(find.text('Cancel'));
      await tester.pumpAndSettle();

      expect(find.byType(ExcelUploadDialog), findsNothing);
    });

    testWidgets('automatically loads active draft on startup for admin department and renders draft canvas', (tester) async {
      final fakeRepo = FakeAdminScheduleDraftRepository();
      fakeRepo.draftToReturn = FakeAdminScheduleDraftRepository._createSampleDraft();

      await tester.pumpWidget(createWidgetUnderTest(repository: fakeRepo));
      await tester.pumpAndSettle();

      expect(fakeRepo.getActiveDraftCallCount, equals(1));
      expect(find.text('OPTIMAL'), findsOneWidget);
      expect(find.text('Dr. Gregory House'), findsOneWidget);
      expect(find.text('Awaiting Generation'), findsNothing);
    });

    testWidgets('when no active draft exists on startup, canvas remains in awaiting generation state', (tester) async {
      final fakeRepo = FakeAdminScheduleDraftRepository();
      fakeRepo.draftToReturn = null;

      await tester.pumpWidget(createWidgetUnderTest(repository: fakeRepo));
      await tester.pumpAndSettle();

      expect(fakeRepo.getActiveDraftCallCount, equals(1));
      expect(find.text('Awaiting Generation'), findsOneWidget);
      expect(find.text('OPTIMAL'), findsNothing);
    });

    testWidgets('tapping Export to Excel on canvas invokes export draft', (tester) async {
      final fakeRepo = FakeAdminScheduleDraftRepository();
      fakeRepo.draftToReturn = FakeAdminScheduleDraftRepository._createSampleDraft();

      await tester.pumpWidget(createWidgetUnderTest(repository: fakeRepo));
      await tester.pumpAndSettle();

      final exportButtonFinder = find.widgetWithText(OutlinedButton, 'Export to Excel');
      expect(exportButtonFinder, findsOneWidget);

      await tester.ensureVisible(exportButtonFinder);
      await tester.pumpAndSettle();

      await tester.tap(exportButtonFinder);
      await tester.pumpAndSettle();

      expect(fakeRepo.exportExcelCallCount, equals(1));
      expect(fakeRepo.capturedDraftId, equals('draft-101'));
      expect(find.text('Schedule exported successfully'), findsOneWidget);
    });

    testWidgets('when active draft fetch fails with an error on startup, an error SnackBar is displayed', (tester) async {
      final fakeRepo = FakeAdminScheduleDraftRepository();
      fakeRepo.errorToThrow = const ApiException(
        type: ApiErrorType.server,
        message: 'Internal server error occurred',
        statusCode: 500,
      );

      await tester.pumpWidget(createWidgetUnderTest(repository: fakeRepo));
      await tester.pumpAndSettle();

      expect(find.text('Internal server error occurred'), findsOneWidget);
    });

    testWidgets('selecting month from popup menu loads draft for selected month', (tester) async {
      final fakeRepo = FakeAdminScheduleDraftRepository();
      fakeRepo.historyToReturn = [
        ScheduleSummary(
          id: 's-100',
          departmentId: 'dept-er',
          targetMonth: '2026-10',
          status: 'published',
          totalDuties: 28,
          solverStatus: 'OPTIMAL',
          sourceFilename: 'oct.xlsx',
          createdAt: DateTime.parse('2026-10-01T00:00:00Z'),
          updatedAt: DateTime.parse('2026-10-01T00:00:00Z'),
        ),
      ];

      await tester.pumpWidget(createWidgetUnderTest(repository: fakeRepo));
      await tester.pumpAndSettle();

      // Open popup menu on Hero Card
      await tester.tap(find.byType(PopupMenuButton<String>));
      await tester.pumpAndSettle();

      // Tap October in popup
      await tester.tap(find.text('2026-10'));
      await tester.pumpAndSettle();

      // getActiveDraft should have been called twice (startup + October selection)
      expect(fakeRepo.getActiveDraftCallCount, equals(2));
    });

    testWidgets('tapping Publish Schedule shows confirmation dialog and publishes on confirm', (tester) async {
      final fakeRepo = FakeAdminScheduleDraftRepository();
      fakeRepo.draftToReturn = FakeAdminScheduleDraftRepository._createSampleDraft();

      await tester.pumpWidget(createWidgetUnderTest(repository: fakeRepo));
      await tester.pumpAndSettle();

      final publishButtonFinder = find.widgetWithText(FilledButton, 'Publish Schedule');
      expect(publishButtonFinder, findsOneWidget);

      await tester.ensureVisible(publishButtonFinder);
      await tester.tap(publishButtonFinder);
      await tester.pumpAndSettle();

      // Dialog should be visible
      expect(find.byType(AlertDialog), findsOneWidget);
      expect(find.text('Publish Schedule'), findsNWidgets(2)); // Button + Dialog Title

      // Tap Cancel first
      await tester.tap(find.text('Cancel'));
      await tester.pumpAndSettle();
      expect(fakeRepo.publishScheduleDraftCallCount, equals(0));

      // Tap Publish button again, then Confirm
      await tester.tap(publishButtonFinder);
      await tester.pumpAndSettle();

      final confirmButtonFinder = find.descendant(
        of: find.byType(AlertDialog),
        matching: find.text('Publish'),
      );
      await tester.tap(confirmButtonFinder);
      await tester.pumpAndSettle();

      expect(fakeRepo.publishScheduleDraftCallCount, equals(1));
      expect(find.textContaining('published successfully'), findsOneWidget);
    });
  });
}
