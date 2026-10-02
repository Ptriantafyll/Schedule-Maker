import 'package:frontend/features/admin/domain/models/schedule_draft.dart';

enum GenerationStatus { idle, solving, success, error }

class ScheduleGenerationState {
  const ScheduleGenerationState({
    this.status = GenerationStatus.idle,
    this.source = '',
    this.errorMessage,
    this.generatedAt,
    this.draft,
    this.isExporting = false,
  });

  final GenerationStatus status;
  final String source;
  final String? errorMessage;
  final DateTime? generatedAt;
  final ScheduleDraft? draft;
  final bool isExporting;

  bool get isGenerated => status == GenerationStatus.success;
  bool get isSolving => status == GenerationStatus.solving;

  ScheduleGenerationState copyWith({
    GenerationStatus? status,
    String? source,
    String? errorMessage,
    DateTime? generatedAt,
    ScheduleDraft? draft,
    bool? isExporting,
    bool clearError = false,
    bool clearDraft = false,
  }) {
    return ScheduleGenerationState(
      status: status ?? this.status,
      source: source ?? this.source,
      errorMessage: clearError ? null : (errorMessage ?? this.errorMessage),
      generatedAt: generatedAt ?? this.generatedAt,
      draft: clearDraft ? null : (draft ?? this.draft),
      isExporting: isExporting ?? this.isExporting,
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
        other.isExporting == isExporting;
  }

  @override
  int get hashCode => Object.hash(
    status,
    source,
    errorMessage,
    generatedAt,
    draft,
    isExporting,
  );
}
