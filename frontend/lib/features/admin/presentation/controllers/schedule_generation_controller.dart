import 'package:file_picker/file_picker.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';

enum GenerationStatus { idle, solving, success, error }

class ScheduleGenerationState {
  const ScheduleGenerationState({
    this.status = GenerationStatus.idle,
    this.source = '',
    this.errorMessage,
    this.generatedAt,
  });

  final GenerationStatus status;
  final String source;
  final String? errorMessage;
  final DateTime? generatedAt;

  bool get isGenerated => status == GenerationStatus.success;
  bool get isSolving => status == GenerationStatus.solving;

  ScheduleGenerationState copyWith({
    GenerationStatus? status,
    String? source,
    String? errorMessage,
    DateTime? generatedAt,
  }) {
    return ScheduleGenerationState(
      status: status ?? this.status,
      source: source ?? this.source,
      errorMessage: errorMessage,
      generatedAt: generatedAt ?? this.generatedAt,
    );
  }
}

abstract class ScheduleRepository {
  Future<void> generateFromRoster({required String month});
  Future<void> generateFromExcel({required PlatformFile file});
}

class DefaultScheduleRepository implements ScheduleRepository {
  @override
  Future<void> generateFromRoster({required String month}) async {
    await Future.delayed(const Duration(milliseconds: 500));
  }

  @override
  Future<void> generateFromExcel({required PlatformFile file}) async {
    await Future.delayed(const Duration(milliseconds: 500));
  }
}

final scheduleRepositoryProvider = Provider<ScheduleRepository>((ref) {
  return DefaultScheduleRepository();
});

class ScheduleGenerationController extends Notifier<ScheduleGenerationState> {
  @override
  ScheduleGenerationState build() => ScheduleGenerationState();

  Future<void> generateFromCurrentRoster({String month = 'November'}) async {
    state = state.copyWith(
      status: GenerationStatus.solving,
      errorMessage: null,
    );

    try {
      final repository = ref.read(scheduleRepositoryProvider);
      await repository.generateFromRoster(month: month);
      state = state.copyWith(
        status: GenerationStatus.success,
        generatedAt: DateTime.now(),
        source: 'Current Department Roster',
      );
    } catch (e) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: e.toString(),
      );
    }
  }

  Future<void> generateFromExcel(PlatformFile file) async {
    state = state.copyWith(
      status: GenerationStatus.solving,
      errorMessage: null,
    );

    try {
      final repository = ref.read(scheduleRepositoryProvider);
      await repository.generateFromExcel(file: file);
      state = state.copyWith(
        status: GenerationStatus.success,
        generatedAt: DateTime.now(),
        source: file.name,
      );
    } catch (e) {
      state = state.copyWith(
        status: GenerationStatus.error,
        errorMessage: e.toString(),
      );
    }
  }

  void reset() {
    state = const ScheduleGenerationState();
  }
}

final scheduleGenerationControllerProvider =
    NotifierProvider<ScheduleGenerationController, ScheduleGenerationState>(
      ScheduleGenerationController.new,
    );
