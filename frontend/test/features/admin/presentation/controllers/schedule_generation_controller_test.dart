import 'dart:typed_data';
// ignore: depend_on_referenced_packages
import 'package:cross_file/cross_file.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/data/repositories/schedule_draft_repository.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';
import 'package:frontend/features/admin/domain/models/schedule_summary.dart';
import 'package:frontend/features/admin/domain/models/target_month_info.dart';
import 'package:frontend/features/admin/presentation/controllers/schedule_generation_controller.dart';

base class FakePlatformFile extends PlatformFile {
  FakePlatformFile({
    required this.name,
    this.fileSize = 1024,
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

class FakeScheduleDraftRepository implements ScheduleDraftRepository {
  int generateFromExcelCallCount = 0;
  int getActiveDraftCallCount = 0;
  int exportExcelCallCount = 0;
  int generateFromRosterCallCount = 0;
  int fetchScheduleHistoryCallCount = 0;
  int fetchTargetMonthInfoCallCount = 0;
  int publishScheduleDraftCallCount = 0;

  PlatformFile? capturedFile;
  String? capturedTargetMonth;
  String? capturedDepartmentId;

  ScheduleDraft? scheduleDraftToReturn;
  List<int> exportBytesToReturn = [1, 2, 3];
  List<ScheduleSummary> historyToReturn = const [];
  TargetMonthInfo targetMonthInfoToReturn = const TargetMonthInfo(nextTargetMonth: '2026-12');
  ScheduleDraft? publishDraftToReturn;
  Exception? exceptionToThrow;

  @override
  Future<ScheduleDraft> generateFromExcel({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
  }) async {
    generateFromExcelCallCount++;
    capturedFile = file;
    capturedTargetMonth = targetMonth;
    capturedDepartmentId = departmentId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return scheduleDraftToReturn ?? _createDummyDraft();
  }

  @override
  Future<ScheduleDraft?> getActiveDraft({
    required String targetMonth,
    String? departmentId,
  }) async {
    getActiveDraftCallCount++;
    capturedTargetMonth = targetMonth;
    capturedDepartmentId = departmentId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return scheduleDraftToReturn;
  }

  String? capturedDraftId;

  @override
  Future<List<int>> exportExcel({
    required String draftId,
  }) async {
    exportExcelCallCount++;
    capturedDraftId = draftId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return exportBytesToReturn;
  }

  @override
  Future<void> generateFromRoster({required String month}) async {
    generateFromRosterCallCount++;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<List<ScheduleSummary>> fetchScheduleHistory({String? departmentId}) async {
    fetchScheduleHistoryCallCount++;
    capturedDepartmentId = departmentId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return historyToReturn;
  }

  @override
  Future<TargetMonthInfo> fetchTargetMonthInfo({String? departmentId}) async {
    fetchTargetMonthInfoCallCount++;
    capturedDepartmentId = departmentId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return targetMonthInfoToReturn;
  }

  @override
  Future<ScheduleDraft> publishScheduleDraft({required String draftId}) async {
    publishScheduleDraftCallCount++;
    capturedDraftId = draftId;
    if (exceptionToThrow != null) throw exceptionToThrow!;
    return publishDraftToReturn ?? scheduleDraftToReturn ?? _createDummyDraft();
  }

  static ScheduleDraft _createDummyDraft() {
    return ScheduleDraft(
      id: 'draft-1',
      departmentId: 'dept-1',
      targetMonth: '2026-11',
      sourceFilename: 'nov.xlsx',
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
  late FakeScheduleDraftRepository fakeRepository;
  late ProviderContainer container;

  setUp(() {
    fakeRepository = FakeScheduleDraftRepository();
    container = ProviderContainer(
      overrides: [
        scheduleDraftRepositoryProvider.overrideWithValue(fakeRepository),
      ],
    );
  });

  tearDown(() {
    container.dispose();
  });

  group('ScheduleGenerationController Tests', () {
    test('initial state is idle with no draft or export active', () {
      final state = container.read(scheduleGenerationControllerProvider);

      expect(state.status, equals(GenerationStatus.idle));
      expect(state.isGenerated, isFalse);
      expect(state.isSolving, isFalse);
      expect(state.isExporting, isFalse);
      expect(state.draft, isNull);
      expect(state.errorMessage, isNull);
      expect(state.source, isEmpty);
      expect(state.generatedAt, isNull);
    });

    test('loadActiveDraft sets success state and assigns draft when draft exists', () async {
      final dummyDraft = FakeScheduleDraftRepository._createDummyDraft();
      fakeRepository.scheduleDraftToReturn = dummyDraft;
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.loadActiveDraft(
        targetMonth: '2026-11',
        departmentId: 'dept-1',
      );

      final state = container.read(scheduleGenerationControllerProvider);
      expect(fakeRepository.getActiveDraftCallCount, equals(1));
      expect(fakeRepository.capturedTargetMonth, equals('2026-11'));
      expect(fakeRepository.capturedDepartmentId, equals('dept-1'));
      expect(state.status, equals(GenerationStatus.success));
      expect(state.isGenerated, isTrue);
      expect(state.draft, equals(dummyDraft));
      expect(state.source, equals('nov.xlsx'));
      expect(state.generatedAt, equals(dummyDraft.createdAt));
    });

    test('loadActiveDraft leaves state idle when no active draft exists', () async {
      fakeRepository.scheduleDraftToReturn = null;
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.loadActiveDraft(
        targetMonth: '2026-11',
        departmentId: 'dept-1',
      );

      final state = container.read(scheduleGenerationControllerProvider);
      expect(fakeRepository.getActiveDraftCallCount, equals(1));
      expect(state.status, equals(GenerationStatus.idle));
      expect(state.isGenerated, isFalse);
      expect(state.draft, isNull);
    });

    test('loadActiveDraft transitions to error when repository throws exception', () async {
      fakeRepository.exceptionToThrow = Exception('Network error');
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.loadActiveDraft(
        targetMonth: '2026-11',
      );

      final state = container.read(scheduleGenerationControllerProvider);
      expect(state.status, equals(GenerationStatus.error));
      expect(state.draft, isNull);
      expect(state.errorMessage, contains('Network error'));
    });

    test('generateFromExcel transitions to success and stores generated draft', () async {
      final dummyDraft = FakeScheduleDraftRepository._createDummyDraft();
      fakeRepository.scheduleDraftToReturn = dummyDraft;
      final controller = container.read(scheduleGenerationControllerProvider.notifier);
      final testFile = FakePlatformFile(name: 'november_schedule.xlsx');

      final future = controller.generateFromExcel(
        testFile,
        targetMonth: '2026-11',
        departmentId: 'dept-1',
      );

      expect(container.read(scheduleGenerationControllerProvider).isSolving, isTrue);

      await future;

      final state = container.read(scheduleGenerationControllerProvider);
      expect(fakeRepository.generateFromExcelCallCount, equals(1));
      expect(fakeRepository.capturedFile?.name, equals('november_schedule.xlsx'));
      expect(fakeRepository.capturedTargetMonth, equals('2026-11'));
      expect(fakeRepository.capturedDepartmentId, equals('dept-1'));
      expect(state.status, equals(GenerationStatus.success));
      expect(state.isGenerated, isTrue);
      expect(state.isSolving, isFalse);
      expect(state.draft, equals(dummyDraft));
      expect(state.source, equals('november_schedule.xlsx'));
      expect(state.generatedAt, isNotNull);
      expect(state.errorMessage, isNull);
    });

    test('generateFromExcel transitions to error state when solver throws exception', () async {
      fakeRepository.exceptionToThrow = Exception('Infeasible shift requirements');
      final controller = container.read(scheduleGenerationControllerProvider.notifier);
      final testFile = FakePlatformFile(name: 'infeasible.xlsx');

      await controller.generateFromExcel(
        testFile,
        targetMonth: '2026-11',
      );

      final state = container.read(scheduleGenerationControllerProvider);
      expect(state.status, equals(GenerationStatus.error));
      expect(state.isGenerated, isFalse);
      expect(state.isSolving, isFalse);
      expect(state.draft, isNull);
      expect(state.errorMessage, contains('Infeasible shift requirements'));
    });

    test('exportCurrentDraft toggles isExporting flag and returns file bytes', () async {
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      final bytesFuture = controller.exportCurrentDraft(
        draftId: 'draft-1',
        targetMonth: '2026-11',
        departmentId: 'dept-1',
      );

      expect(container.read(scheduleGenerationControllerProvider).isExporting, isTrue);

      final bytes = await bytesFuture;

      expect(bytes, equals([1, 2, 3]));
      expect(fakeRepository.exportExcelCallCount, equals(1));
      expect(fakeRepository.capturedDraftId, equals('draft-1'));
      expect(container.read(scheduleGenerationControllerProvider).isExporting, isFalse);
    });

    test('exportCurrentDraft resets isExporting to false on failure', () async {
      fakeRepository.exceptionToThrow = Exception('Export failed');
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      final bytes = await controller.exportCurrentDraft(
        draftId: 'draft-1',
        targetMonth: '2026-11',
      );

      expect(bytes, isNull);
      final state = container.read(scheduleGenerationControllerProvider);
      expect(state.isExporting, isFalse);
      expect(state.errorMessage, contains('Export failed'));
    });

    test('generateFromCurrentRoster transitions to success', () async {
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.generateFromCurrentRoster(month: 'November');

      final state = container.read(scheduleGenerationControllerProvider);
      expect(fakeRepository.generateFromRosterCallCount, equals(1));
      expect(state.status, equals(GenerationStatus.success));
      expect(state.isGenerated, isTrue);
    });

    test('reset clears state and draft back to idle', () async {
      final dummyDraft = FakeScheduleDraftRepository._createDummyDraft();
      fakeRepository.scheduleDraftToReturn = dummyDraft;
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.loadActiveDraft(targetMonth: '2026-11');
      expect(container.read(scheduleGenerationControllerProvider).isGenerated, isTrue);

      controller.reset();

      final state = container.read(scheduleGenerationControllerProvider);
      expect(state.status, equals(GenerationStatus.idle));
      expect(state.isGenerated, isFalse);
      expect(state.draft, isNull);
      expect(state.errorMessage, isNull);
      expect(state.source, isEmpty);
    });

    final sampleSummary = ScheduleSummary(
      id: 'summary-1',
      departmentId: 'dept-1',
      targetMonth: '2026-11',
      sourceFilename: 'nov.xlsx',
      totalDuties: 28,
      solverStatus: 'OPTIMAL',
      status: 'published',
      createdAt: DateTime.parse('2026-11-01T08:00:00.000Z'),
      updatedAt: DateTime.parse('2026-11-01T08:30:00.000Z'),
    );

    test('initializeDashboard fetches targetMonthInfo and history, sets selectedMonth, and loads active draft', () async {
      final dummyDraft = FakeScheduleDraftRepository._createDummyDraft();
      fakeRepository.targetMonthInfoToReturn = const TargetMonthInfo(
        nextTargetMonth: '2026-12',
        lastPublishedMonth: '2026-11',
      );
      fakeRepository.historyToReturn = [sampleSummary];
      fakeRepository.scheduleDraftToReturn = dummyDraft;

      final controller = container.read(scheduleGenerationControllerProvider.notifier);
      await controller.initializeDashboard(departmentId: 'dept-1');

      final state = container.read(scheduleGenerationControllerProvider);
      expect(fakeRepository.fetchTargetMonthInfoCallCount, equals(1));
      expect(fakeRepository.fetchScheduleHistoryCallCount, equals(1));
      expect(fakeRepository.getActiveDraftCallCount, equals(1));
      expect(fakeRepository.capturedTargetMonth, equals('2026-12'));
      expect(fakeRepository.capturedDepartmentId, equals('dept-1'));
      expect(state.selectedMonth, equals('2026-12'));
      expect(state.targetMonthInfo?.nextTargetMonth, equals('2026-12'));
      expect(state.scheduleHistory, equals([sampleSummary]));
      expect(state.draft, equals(dummyDraft));
      expect(state.isGenerated, isTrue);
    });

    test('initializeDashboard sets error state when repository throws', () async {
      fakeRepository.exceptionToThrow = Exception('Dashboard init failed');
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.initializeDashboard(departmentId: 'dept-1');

      final state = container.read(scheduleGenerationControllerProvider);
      expect(state.status, equals(GenerationStatus.error));
      expect(state.errorMessage, contains('Dashboard init failed'));
    });

    test('selectMonth updates selectedMonth and loads active draft for that month', () async {
      final dummyDraft = FakeScheduleDraftRepository._createDummyDraft();
      fakeRepository.scheduleDraftToReturn = dummyDraft;
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.selectMonth('2027-01', departmentId: 'dept-1');

      final state = container.read(scheduleGenerationControllerProvider);
      expect(fakeRepository.getActiveDraftCallCount, equals(1));
      expect(fakeRepository.capturedTargetMonth, equals('2027-01'));
      expect(state.selectedMonth, equals('2027-01'));
      expect(state.draft, equals(dummyDraft));
    });

    test('generateFromExcel uses state.selectedMonth when targetMonth is not explicitly passed', () async {
      final controller = container.read(scheduleGenerationControllerProvider.notifier);
      await controller.selectMonth('2027-02');

      final testFile = FakePlatformFile(name: 'feb.xlsx');
      await controller.generateFromExcel(testFile);

      expect(fakeRepository.capturedTargetMonth, equals('2027-02'));
    });

    test('publishCurrentDraft publishes active draft and refreshes target info and history', () async {
      final dummyDraft = FakeScheduleDraftRepository._createDummyDraft();
      fakeRepository.scheduleDraftToReturn = dummyDraft;
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.loadActiveDraft(targetMonth: '2026-11');

      final publishedDraft = ScheduleDraft(
        id: dummyDraft.id,
        departmentId: dummyDraft.departmentId,
        targetMonth: dummyDraft.targetMonth,
        sourceFilename: dummyDraft.sourceFilename,
        totalDuties: dummyDraft.totalDuties,
        solverStatus: dummyDraft.solverStatus,
        status: 'published',
        assignments: dummyDraft.assignments,
        unavailabilities: dummyDraft.unavailabilities,
        createdAt: dummyDraft.createdAt,
        updatedAt: DateTime.parse('2026-11-01T09:00:00.000Z'),
      );
      fakeRepository.publishDraftToReturn = publishedDraft;

      await controller.publishCurrentDraft(departmentId: 'dept-1');

      final state = container.read(scheduleGenerationControllerProvider);
      expect(fakeRepository.publishScheduleDraftCallCount, equals(1));
      expect(fakeRepository.capturedDraftId, equals('draft-1'));
      expect(state.draft?.status, equals('published'));
      expect(state.isPublished, isTrue);
      expect(state.isPublishing, isFalse);
      expect(fakeRepository.fetchTargetMonthInfoCallCount, equals(1));
      expect(fakeRepository.fetchScheduleHistoryCallCount, equals(1));
    });

    test('publishCurrentDraft sets error when no active draft is loaded', () async {
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.publishCurrentDraft();

      final state = container.read(scheduleGenerationControllerProvider);
      expect(state.status, equals(GenerationStatus.error));
      expect(state.errorMessage, equals('No active schedule draft to publish.'));
      expect(fakeRepository.publishScheduleDraftCallCount, equals(0));
    });

    test('publishCurrentDraft resets isPublishing and sets error when repository throws', () async {
      final dummyDraft = FakeScheduleDraftRepository._createDummyDraft();
      fakeRepository.scheduleDraftToReturn = dummyDraft;
      final controller = container.read(scheduleGenerationControllerProvider.notifier);
      await controller.loadActiveDraft(targetMonth: '2026-11');

      fakeRepository.exceptionToThrow = Exception('Publish conflict');

      await controller.publishCurrentDraft();

      final state = container.read(scheduleGenerationControllerProvider);
      expect(state.isPublishing, isFalse);
      expect(state.status, equals(GenerationStatus.error));
      expect(state.errorMessage, contains('Publish conflict'));
    });
  });
}
