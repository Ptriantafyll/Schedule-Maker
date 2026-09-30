import 'dart:typed_data';
// ignore: depend_on_referenced_packages
import 'package:cross_file/cross_file.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
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

class FakeScheduleRepository implements ScheduleRepository {
  int generateFromRosterCallCount = 0;
  int generateFromExcelCallCount = 0;

  String? capturedMonth;
  PlatformFile? capturedFile;

  Exception? exceptionToThrow;

  @override
  Future<void> generateFromRoster({required String month}) async {
    generateFromRosterCallCount++;
    capturedMonth = month;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }

  @override
  Future<void> generateFromExcel({required PlatformFile file}) async {
    generateFromExcelCallCount++;
    capturedFile = file;
    if (exceptionToThrow != null) throw exceptionToThrow!;
  }
}

void main() {
  late FakeScheduleRepository fakeRepository;
  late ProviderContainer container;

  setUp(() {
    fakeRepository = FakeScheduleRepository();
    container = ProviderContainer(
      overrides: [
        scheduleRepositoryProvider.overrideWithValue(fakeRepository),
      ],
    );
  });

  tearDown(() {
    container.dispose();
  });

  group('ScheduleGenerationController Tests', () {
    test('initial state is idle and not generated', () {
      final state = container.read(scheduleGenerationControllerProvider);

      expect(state.status, equals(GenerationStatus.idle));
      expect(state.isGenerated, isFalse);
      expect(state.isSolving, isFalse);
      expect(state.errorMessage, isNull);
      expect(state.source, isEmpty);
      expect(state.generatedAt, isNull);
    });

    test('generateFromCurrentRoster successfully transitions to success state', () async {
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      final future = controller.generateFromCurrentRoster(month: 'November');

      // Immediate state while solving
      expect(container.read(scheduleGenerationControllerProvider).isSolving, isTrue);

      await future;

      final state = container.read(scheduleGenerationControllerProvider);
      expect(fakeRepository.generateFromRosterCallCount, equals(1));
      expect(fakeRepository.capturedMonth, equals('November'));
      expect(state.status, equals(GenerationStatus.success));
      expect(state.isGenerated, isTrue);
      expect(state.isSolving, isFalse);
      expect(state.source, equals('Current Department Roster'));
      expect(state.generatedAt, isNotNull);
      expect(state.errorMessage, isNull);
    });

    test('generateFromExcel successfully transitions to success state with filename as source', () async {
      final controller = container.read(scheduleGenerationControllerProvider.notifier);
      final testFile = FakePlatformFile(name: 'november_schedule.xlsx');

      final future = controller.generateFromExcel(testFile);

      // Immediate state while solving
      expect(container.read(scheduleGenerationControllerProvider).isSolving, isTrue);

      await future;

      final state = container.read(scheduleGenerationControllerProvider);
      expect(fakeRepository.generateFromExcelCallCount, equals(1));
      expect(fakeRepository.capturedFile?.name, equals('november_schedule.xlsx'));
      expect(state.status, equals(GenerationStatus.success));
      expect(state.isGenerated, isTrue);
      expect(state.isSolving, isFalse);
      expect(state.source, equals('november_schedule.xlsx'));
      expect(state.generatedAt, isNotNull);
      expect(state.errorMessage, isNull);
    });

    test('captures error message and sets error status when solver throws exception', () async {
      fakeRepository.exceptionToThrow = Exception('Infeasible shift requirements');
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.generateFromCurrentRoster(month: 'December');

      final state = container.read(scheduleGenerationControllerProvider);
      expect(state.status, equals(GenerationStatus.error));
      expect(state.isGenerated, isFalse);
      expect(state.isSolving, isFalse);
      expect(state.errorMessage, contains('Infeasible shift requirements'));
    });

    test('reset clears generated state and returns to idle', () async {
      final controller = container.read(scheduleGenerationControllerProvider.notifier);

      await controller.generateFromCurrentRoster(month: 'November');
      expect(container.read(scheduleGenerationControllerProvider).isGenerated, isTrue);

      controller.reset();

      final state = container.read(scheduleGenerationControllerProvider);
      expect(state.status, equals(GenerationStatus.idle));
      expect(state.isGenerated, isFalse);
      expect(state.errorMessage, isNull);
      expect(state.source, isEmpty);
    });
  });
}
