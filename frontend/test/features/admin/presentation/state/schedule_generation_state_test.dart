import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';
import 'package:frontend/features/admin/domain/models/schedule_summary.dart';
import 'package:frontend/features/admin/domain/models/target_month_info.dart';
import 'package:frontend/features/admin/presentation/state/schedule_generation_state.dart';

void main() {
  final sampleDraft = ScheduleDraft(
    id: 'draft-1',
    departmentId: 'dept-1',
    targetMonth: '2026-11',
    sourceFilename: 'nov.xlsx',
    totalDuties: 28,
    solverStatus: 'OPTIMAL',
    status: 'draft',
    assignments: const [],
    unavailabilities: const {},
    createdAt: DateTime.parse('2026-11-01T08:00:00.000Z'),
    updatedAt: DateTime.parse('2026-11-01T08:00:00.000Z'),
  );

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

  const sampleTargetMonthInfo = TargetMonthInfo(
    nextTargetMonth: '2026-12',
    lastPublishedMonth: '2026-11',
  );

  group('ScheduleGenerationState Tests', () {
    test('initial state has correct default values', () {
      const state = ScheduleGenerationState();

      expect(state.status, equals(GenerationStatus.idle));
      expect(state.source, isEmpty);
      expect(state.errorMessage, isNull);
      expect(state.generatedAt, isNull);
      expect(state.draft, isNull);
      expect(state.isExporting, isFalse);
      expect(state.isGenerated, isFalse);
      expect(state.isSolving, isFalse);
      expect(state.selectedMonth, isEmpty);
      expect(state.targetMonthInfo, isNull);
      expect(state.scheduleHistory, isEmpty);
      expect(state.isPublishing, isFalse);
      expect(state.isPublished, isFalse);
    });

    test('computed getters reflect current status accurately', () {
      const solvingState = ScheduleGenerationState(
        status: GenerationStatus.solving,
      );
      expect(solvingState.isSolving, isTrue);
      expect(solvingState.isGenerated, isFalse);

      const successState = ScheduleGenerationState(
        status: GenerationStatus.success,
      );
      expect(successState.isSolving, isFalse);
      expect(successState.isGenerated, isTrue);

      final publishedDraft = ScheduleDraft(
        id: 'draft-1',
        departmentId: 'dept-1',
        targetMonth: '2026-11',
        sourceFilename: 'nov.xlsx',
        totalDuties: 28,
        solverStatus: 'OPTIMAL',
        status: 'published',
        assignments: const [],
        unavailabilities: const {},
        createdAt: DateTime.parse('2026-11-01T08:00:00.000Z'),
        updatedAt: DateTime.parse('2026-11-01T08:00:00.000Z'),
      );
      final publishedState = ScheduleGenerationState(draft: publishedDraft);
      expect(publishedState.isPublished, isTrue);
    });

    test('copyWith updates specified fields and preserves others', () {
      const initial = ScheduleGenerationState();

      final updated = initial.copyWith(
        status: GenerationStatus.success,
        source: 'nov.xlsx',
        draft: sampleDraft,
        isExporting: true,
        selectedMonth: '2026-12',
        targetMonthInfo: sampleTargetMonthInfo,
        scheduleHistory: [sampleSummary],
        isPublishing: true,
      );

      expect(updated.status, equals(GenerationStatus.success));
      expect(updated.source, equals('nov.xlsx'));
      expect(updated.draft, equals(sampleDraft));
      expect(updated.isExporting, isTrue);
      expect(updated.selectedMonth, equals('2026-12'));
      expect(updated.targetMonthInfo, equals(sampleTargetMonthInfo));
      expect(updated.scheduleHistory, equals([sampleSummary]));
      expect(updated.isPublishing, isTrue);
      expect(updated.errorMessage, isNull);
    });

    test('copyWith preserves existing errorMessage unless explicitly cleared', () {
      const initial = ScheduleGenerationState(
        errorMessage: 'Something failed',
      );

      final updated = initial.copyWith(isExporting: false);
      expect(updated.errorMessage, equals('Something failed'));

      final cleared = initial.copyWith(clearError: true);
      expect(cleared.errorMessage, isNull);
    });

    test('copyWith clearDraft flag resets draft to null', () {
      final stateWithDraft = ScheduleGenerationState(
        draft: sampleDraft,
      );
      expect(stateWithDraft.draft, isNotNull);

      final cleared = stateWithDraft.copyWith(clearDraft: true);
      expect(cleared.draft, isNull);
    });

    test('equality and hashCode compare field values correctly', () {
      final stateA = ScheduleGenerationState(
        status: GenerationStatus.success,
        draft: sampleDraft,
        selectedMonth: '2026-12',
        targetMonthInfo: sampleTargetMonthInfo,
        scheduleHistory: [sampleSummary],
        isPublishing: false,
      );
      final stateB = ScheduleGenerationState(
        status: GenerationStatus.success,
        draft: sampleDraft,
        selectedMonth: '2026-12',
        targetMonthInfo: sampleTargetMonthInfo,
        scheduleHistory: [sampleSummary],
        isPublishing: false,
      );
      const stateC = ScheduleGenerationState();

      expect(stateA, equals(stateB));
      expect(stateA.hashCode, equals(stateB.hashCode));
      expect(stateA, isNot(equals(stateC)));
    });
  });
}
