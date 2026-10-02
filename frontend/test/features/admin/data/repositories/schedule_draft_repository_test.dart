import 'dart:typed_data';
// ignore: depend_on_referenced_packages
import 'package:cross_file/cross_file.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/data/datasources/schedule_draft_remote_data_source.dart';
import 'package:frontend/features/admin/data/repositories/schedule_draft_repository.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';

base class FakePlatformFile extends PlatformFile {
  FakePlatformFile({
    required this.name,
    required this.fileSize,
  });

  @override
  final String name;

  final int fileSize;

  @override
  Future<Uint8List> readAsBytes() async => Uint8List(0);

  @override
  int? lengthSync() => fileSize;

  @override
  Future<int?> length() async => fileSize;

  @override
  Stream<Uint8List> readAsByteStream() => Stream.value(Uint8List(0));

  @override
  Uri get uri => Uri.file(name);

  @override
  XFile get xFile => XFile(name);
}

class FakeScheduleDraftRemoteDataSource implements ScheduleDraftRemoteDataSource {
  PlatformFile? capturedFile;
  String? capturedTargetMonth;
  String? capturedDepartmentId;

  ScheduleDraft? draftToReturn;
  List<int>? bytesToReturn;
  Exception? exceptionToThrow;

  @override
  Future<ScheduleDraft> generateFromExcel({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
  }) async {
    capturedFile = file;
    capturedTargetMonth = targetMonth;
    capturedDepartmentId = departmentId;

    if (exceptionToThrow != null) throw exceptionToThrow!;
    return draftToReturn!;
  }

  @override
  Future<ScheduleDraft?> getActiveDraft({
    required String targetMonth,
    String? departmentId,
  }) async {
    capturedTargetMonth = targetMonth;
    capturedDepartmentId = departmentId;

    if (exceptionToThrow != null) throw exceptionToThrow!;
    return draftToReturn;
  }

  @override
  Future<List<int>> exportExcel({
    required String targetMonth,
    String? departmentId,
  }) async {
    capturedTargetMonth = targetMonth;
    capturedDepartmentId = departmentId;

    if (exceptionToThrow != null) throw exceptionToThrow!;
    return bytesToReturn ?? const <int>[];
  }
}

void main() {
  late FakeScheduleDraftRemoteDataSource fakeRemoteDataSource;
  late ScheduleDraftRepository repository;

  final sampleDraft = ScheduleDraft(
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

  final testFile = FakePlatformFile(name: 'nov.xlsx', fileSize: 500);

  setUp(() {
    fakeRemoteDataSource = FakeScheduleDraftRemoteDataSource();
    repository = ScheduleDraftRepositoryImpl(remoteDataSource: fakeRemoteDataSource);
  });

  group('ScheduleDraftRepositoryImpl.generateFromExcel', () {
    test('delegates to remoteDataSource and passes departmentId', () async {
      fakeRemoteDataSource.draftToReturn = sampleDraft;

      final result = await repository.generateFromExcel(
        file: testFile,
        targetMonth: '2026-11',
        departmentId: 'dept-1',
      );

      expect(fakeRemoteDataSource.capturedFile?.name, equals('nov.xlsx'));
      expect(fakeRemoteDataSource.capturedTargetMonth, equals('2026-11'));
      expect(fakeRemoteDataSource.capturedDepartmentId, equals('dept-1'));
      expect(result, equals(sampleDraft));
    });

    test('propagates exception when remoteDataSource throws', () async {
      fakeRemoteDataSource.exceptionToThrow = Exception('Network error');

      expect(
        () => repository.generateFromExcel(file: testFile, targetMonth: '2026-11'),
        throwsException,
      );
    });
  });

  group('ScheduleDraftRepositoryImpl.getActiveDraft', () {
    test('delegates to remoteDataSource and passes departmentId', () async {
      fakeRemoteDataSource.draftToReturn = sampleDraft;

      final result = await repository.getActiveDraft(
        targetMonth: '2026-11',
        departmentId: 'dept-1',
      );

      expect(fakeRemoteDataSource.capturedTargetMonth, equals('2026-11'));
      expect(fakeRemoteDataSource.capturedDepartmentId, equals('dept-1'));
      expect(result, equals(sampleDraft));
    });

    test('returns null when remoteDataSource returns null', () async {
      fakeRemoteDataSource.draftToReturn = null;

      final result = await repository.getActiveDraft(targetMonth: '2026-11');

      expect(result, isNull);
    });
  });

  group('ScheduleDraftRepositoryImpl.exportExcel', () {
    test('delegates to remoteDataSource and passes departmentId', () async {
      fakeRemoteDataSource.bytesToReturn = [1, 2, 3];

      final result = await repository.exportExcel(
        targetMonth: '2026-11',
        departmentId: 'dept-1',
      );

      expect(fakeRemoteDataSource.capturedTargetMonth, equals('2026-11'));
      expect(fakeRemoteDataSource.capturedDepartmentId, equals('dept-1'));
      expect(result, equals([1, 2, 3]));
    });
  });

  group('ScheduleDraftRepositoryImpl.generateFromRoster', () {
    test('completes successfully', () async {
      await expectLater(
        repository.generateFromRoster(month: 'November'),
        completes,
      );
    });
  });
}
