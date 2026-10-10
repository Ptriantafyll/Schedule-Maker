import 'package:flutter/foundation.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';
import 'package:frontend/features/admin/domain/models/schedule_summary.dart';
import 'package:frontend/features/admin/domain/models/target_month_info.dart';

enum GenerationStatus { idle, solving, success, error }

class ScheduleGenerationState {
  const ScheduleGenerationState({
    this.status = GenerationStatus.idle,
    this.source = '',
    this.errorMessage,
    this.generatedAt,
    this.draft,
    this.isExporting = false,
    this.selectedMonth = '',
    this.targetMonthInfo,
    this.scheduleHistory = const [],
    this.isPublishing = false,
  });

  final GenerationStatus status;
  final String source;
  final String? errorMessage;
  final DateTime? generatedAt;
  final ScheduleDraft? draft;
  final bool isExporting;
  final String selectedMonth;
  final TargetMonthInfo? targetMonthInfo;
  final List<ScheduleSummary> scheduleHistory;
  final bool isPublishing;

  bool get isGenerated => status == GenerationStatus.success;
  bool get isSolving => status == GenerationStatus.solving;
  bool get isPublished => draft?.isPublished ?? false;

  ScheduleGenerationState copyWith({
    GenerationStatus? status,
    String? source,
    String? errorMessage,
    DateTime? generatedAt,
    ScheduleDraft? draft,
    bool? isExporting,
    bool clearError = false,
    bool clearDraft = false,
    String? selectedMonth,
    TargetMonthInfo? targetMonthInfo,
    List<ScheduleSummary>? scheduleHistory,
    bool? isPublishing,
  }) {
    return ScheduleGenerationState(
      status: status ?? this.status,
      source: source ?? this.source,
      errorMessage: clearError ? null : (errorMessage ?? this.errorMessage),
      generatedAt: generatedAt ?? this.generatedAt,
      draft: clearDraft ? null : (draft ?? this.draft),
      isExporting: isExporting ?? this.isExporting,
      selectedMonth: selectedMonth ?? this.selectedMonth,
      targetMonthInfo: targetMonthInfo ?? this.targetMonthInfo,
      scheduleHistory: scheduleHistory ?? this.scheduleHistory,
      isPublishing: isPublishing ?? this.isPublishing,
    );
  }

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is ScheduleGenerationState &&
        other.status == status &&
        other.source == source &&
        other.errorMessage == errorMessage &&
        other.generatedAt == generatedAt &&
        other.draft == draft &&
        other.isExporting == isExporting &&
        other.selectedMonth == selectedMonth &&
        other.targetMonthInfo == targetMonthInfo &&
        listEquals(other.scheduleHistory, scheduleHistory) &&
        other.isPublishing == isPublishing;
  }

  @override
  int get hashCode => Object.hash(
    status,
    source,
    errorMessage,
    generatedAt,
    draft,
    isExporting,
    selectedMonth,
    targetMonthInfo,
    Object.hashAll(scheduleHistory),
    isPublishing,
  );
}
