import 'package:file_picker/file_picker.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/data/repositories/schedule_draft_repository.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';
import 'package:frontend/features/admin/presentation/state/schedule_generation_state.dart';

export 'package:frontend/features/admin/presentation/state/schedule_generation_state.dart';

class ScheduleGenerationController extends Notifier<ScheduleGenerationState> {
  @override
  ScheduleGenerationState build() => const ScheduleGenerationState();

  String _extractErrorMessage(Object error) {
    if (error is ApiException) {
      return error.message;
    }
    return error.toString();
  }

  Future<void> loadActiveDraft({
    required String targetMonth,
    String? departmentId,
  }) async {
    state = state.copyWith(clearError: true);

    try {
      final repository = ref.read(scheduleDraftRepositoryProvider);
      final draft = await repository.getActiveDraft(
        targetMonth: targetMonth,
        departmentId: departmentId,
      );
      _applyLoadedDraft(draft);
    } catch (e) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: _extractErrorMessage(e),
        clearDraft: true,
      );
    }
  }

  void _applyLoadedDraft(ScheduleDraft? draft) {
    if (draft == null) {
      state = state.copyWith(status: GenerationStatus.idle, clearDraft: true);
      return;
    }

    state = state.copyWith(
      status: GenerationStatus.success,
      draft: draft,
      source: draft.sourceFilename.isNotEmpty
          ? draft.sourceFilename
          : 'Active Draft',
      generatedAt: draft.createdAt,
    );
  }

  Future<void> generateFromExcel(
    PlatformFile file, {
    String? targetMonth,
    String? departmentId,
  }) async {
    state = state.copyWith(status: GenerationStatus.solving, clearError: true);

    final effectiveTargetMonth = (targetMonth != null && targetMonth.isNotEmpty)
        ? targetMonth
        : (state.selectedMonth.isNotEmpty ? state.selectedMonth : '2026-11');

    try {
      final repository = ref.read(scheduleDraftRepositoryProvider);
      final draft = await repository.generateFromExcel(
        file: file,
        targetMonth: effectiveTargetMonth,
        departmentId: departmentId,
      );
      state = state.copyWith(
        status: GenerationStatus.success,
        draft: draft,
        source: file.name,
        generatedAt: draft.createdAt,
      );
    } catch (e) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: _extractErrorMessage(e),
        clearDraft: true,
      );
    }
  }

  Future<List<int>?> exportCurrentDraft({
    String? draftId,
    required String targetMonth,
    String? departmentId,
  }) async {
    final effectiveDraftId = draftId ?? state.draft?.id;
    if (effectiveDraftId == null) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: 'No active schedule draft to export.',
      );
      return null;
    }

    state = state.copyWith(isExporting: true, clearError: true);

    try {
      final repository = ref.read(scheduleDraftRepositoryProvider);
      return await repository.exportExcel(draftId: effectiveDraftId);
    } catch (e) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: _extractErrorMessage(e),
      );
      return null;
    } finally {
      state = state.copyWith(isExporting: false);
    }
  }

  Future<void> generateFromCurrentRoster({String month = 'November'}) async {
    state = state.copyWith(status: GenerationStatus.solving, clearError: true);

    try {
      final repository = ref.read(scheduleDraftRepositoryProvider);
      await repository.generateFromRoster(month: month);
      state = state.copyWith(
        status: GenerationStatus.success,
        generatedAt: DateTime.now(),
        source: 'Current Department Roster',
      );
    } catch (e) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: _extractErrorMessage(e),
      );
    }
  }

  Future<void> initializeDashboard({String? departmentId}) async {
    state = state.copyWith(clearError: true);
    try {
      final repository = ref.read(scheduleDraftRepositoryProvider);
      final targetMonthInfo = await repository.fetchTargetMonthInfo(
        departmentId: departmentId,
      );
      final history = await repository.fetchScheduleHistory(
        departmentId: departmentId,
      );

      state = state.copyWith(
        targetMonthInfo: targetMonthInfo,
        selectedMonth: targetMonthInfo.nextTargetMonth,
        scheduleHistory: history,
      );

      await loadActiveDraft(
        targetMonth: targetMonthInfo.nextTargetMonth,
        departmentId: departmentId,
      );
    } catch (e) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: _extractErrorMessage(e),
      );
    }
  }

  Future<void> publishCurrentDraft({String? departmentId}) async {
    final draftId = state.draft?.id;
    if (draftId == null) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: 'No active schedule draft to publish.',
      );
      return;
    }

    state = state.copyWith(isPublishing: true, clearError: true);

    try {
      final repository = ref.read(scheduleDraftRepositoryProvider);
      final publishedDraft = await repository.publishScheduleDraft(
        draftId: draftId,
      );
      state = state.copyWith(draft: publishedDraft);

      final targetMonthInfo = await repository.fetchTargetMonthInfo(
        departmentId: departmentId,
      );
      final history = await repository.fetchScheduleHistory(
        departmentId: departmentId,
      );
      state = state.copyWith(
        targetMonthInfo: targetMonthInfo,
        scheduleHistory: history,
      );
    } catch (e) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: _extractErrorMessage(e),
      );
    } finally {
      state = state.copyWith(isPublishing: false);
    }
  }

  Future<void> selectMonth(String targetMonth, {String? departmentId}) async {
    state = state.copyWith(selectedMonth: targetMonth);
    await loadActiveDraft(targetMonth: targetMonth, departmentId: departmentId);
  }

  void reset() {
    state = const ScheduleGenerationState();
  }
}

final scheduleGenerationControllerProvider =
    NotifierProvider<ScheduleGenerationController, ScheduleGenerationState>(
      ScheduleGenerationController.new,
    );
