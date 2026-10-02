import 'dart:typed_data';
// ignore: depend_on_referenced_packages
import 'package:cross_file/cross_file.dart';
import 'package:dio/dio.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/data/datasources/schedule_draft_remote_data_source.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';

base class FakePlatformFile extends PlatformFile {
  FakePlatformFile({
    required this.name,
    required this.fileSize,
    this.fileBytes,
  });

  @override
  final String name;

  final int fileSize;
  final Uint8List? fileBytes;

  @override
  Future<Uint8List> readAsBytes() async => fileBytes ?? Uint8List(0);

  @override
  int? lengthSync() => fileSize;

  @override
  Future<int?> length() async => fileSize;

  @override
  Stream<Uint8List> readAsByteStream() => Stream.value(fileBytes ?? Uint8List(0));

  @override
  Uri get uri => Uri.file(name);

  @override
  XFile get xFile => XFile(name);
}

class FakeApiClient implements ApiClient {
  String? capturedPath;
  Object? capturedData;
  Map<String, dynamic>? capturedQueryParams;
  String? capturedContentType;
  bool? capturedRequiresAuth;

  Response<dynamic>? responseToReturn;
  Exception? exceptionToThrow;

  @override
  Future<Response<T>> post<T>(
    String path, {
    Object? data,
    Map<String, dynamic>? queryParameters,
    String? contentType,
    bool requiresAuth = true,
  }) async {
    capturedPath = path;
    capturedData = data;
    capturedQueryParams = queryParameters;
    capturedContentType = contentType;
    capturedRequiresAuth = requiresAuth;

    if (exceptionToThrow != null) {
      throw exceptionToThrow!;
    }

    return responseToReturn as Response<T>;
  }

  @override
  Future<Response<T>> get<T>(
    String path, {
    Map<String, dynamic>? queryParameters,
    bool requiresAuth = true,
  }) async {
    capturedPath = path;
    capturedQueryParams = queryParameters;
    capturedRequiresAuth = requiresAuth;

    if (exceptionToThrow != null) {
      throw exceptionToThrow!;
    }

    return responseToReturn as Response<T>;
  }
}

void main() {
  late FakeApiClient fakeApiClient;
  late ScheduleDraftRemoteDataSource dataSource;

  final sampleDraftPayload = <String, dynamic>{
    'id': 'draft-uuid-1',
    'department_id': 'dept-uuid-1',
    'target_month': '2026-11',
    'source_filename': 'nov_roster.xlsx',
    'total_duties': 28,
    'solver_status': 'OPTIMAL',
    'status': 'draft',
    'assignments': [
      {
        'date': '2026-11-01',
        'day_name': 'Sunday',
        'doctor_name': 'Dr. Gregory House',
        'doctor_email': 'house@hospital.org',
        'position': 'ER',
        'shift': 'Night',
      }
    ],
    'unavailabilities': {
      'Dr. Gregory House': [3, 14, 22],
    },
    'created_at': '2026-11-01T08:00:00.000Z',
    'updated_at': '2026-11-01T08:05:00.000Z',
  };

  final testFile = FakePlatformFile(
    name: 'nov_roster.xlsx',
    fileSize: 1024,
    fileBytes: Uint8List.fromList([1, 2, 3, 4]),
  );

  setUp(() {
    fakeApiClient = FakeApiClient();
    dataSource = ScheduleDraftRemoteDataSource(fakeApiClient);
  });

  group('ScheduleDraftRemoteDataSource.generateFromExcel', () {
    test('sends multipart POST to /api/v1/schedules/generate-from-excel and returns ScheduleDraft', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/generate-from-excel'),
        statusCode: 201,
        data: sampleDraftPayload,
      );

      final result = await dataSource.generateFromExcel(
        file: testFile,
        targetMonth: '2026-11',
        departmentId: 'dept-uuid-1',
      );

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules/generate-from-excel'));
      expect(fakeApiClient.capturedContentType, equals('multipart/form-data'));
      expect(fakeApiClient.capturedData, isA<FormData>());

      final formData = fakeApiClient.capturedData as FormData;
      expect(formData.fields.any((f) => f.key == 'target_month' && f.value == '2026-11'), isTrue);
      expect(formData.fields.any((f) => f.key == 'department_id' && f.value == 'dept-uuid-1'), isTrue);
      expect(formData.files.any((f) => f.key == 'file' && f.value.filename == 'nov_roster.xlsx'), isTrue);

      expect(result, isA<ScheduleDraft>());
      expect(result.id, equals('draft-uuid-1'));
      expect(result.assignments.length, equals(1));
    });

    test('propagates ApiException when solver endpoint fails', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.validation,
        statusCode: 422,
        message: 'Solver infeasible: conflicting shifts',
      );

      expect(
        () => dataSource.generateFromExcel(
          file: testFile,
          targetMonth: '2026-11',
        ),
        throwsA(isA<ApiException>()),
      );
    });
  });

  group('ScheduleDraftRemoteDataSource.getActiveDraft', () {
    test('sends GET to /api/v1/schedules/draft with query parameters and returns ScheduleDraft', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/draft'),
        statusCode: 200,
        data: sampleDraftPayload,
      );

      final result = await dataSource.getActiveDraft(
        targetMonth: '2026-11',
        departmentId: 'dept-uuid-1',
      );

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules/draft'));
      expect(fakeApiClient.capturedQueryParams?['target_month'], equals('2026-11'));
      expect(fakeApiClient.capturedQueryParams?['department_id'], equals('dept-uuid-1'));
      expect(result, isA<ScheduleDraft>());
      expect(result?.id, equals('draft-uuid-1'));
    });

    test('returns null when draft does not exist (HTTP 404)', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.notFound,
        statusCode: 404,
        message: 'No draft found for this department and month',
      );

      final result = await dataSource.getActiveDraft(
        targetMonth: '2026-11',
      );

      expect(result, isNull);
    });

    test('propagates other ApiException errors (e.g. HTTP 500)', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.server,
        statusCode: 500,
        message: 'Internal server error',
      );

      expect(
        () => dataSource.getActiveDraft(targetMonth: '2026-11'),
        throwsA(isA<ApiException>()),
      );
    });
  });

  group('ScheduleDraftRemoteDataSource.exportExcel', () {
    test('sends GET to /api/v1/schedules/export-excel and returns bytes', () async {
      final sampleBytes = [10, 20, 30, 40];
      fakeApiClient.responseToReturn = Response<List<int>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/export-excel'),
        statusCode: 200,
        data: sampleBytes,
      );

      final result = await dataSource.exportExcel(
        targetMonth: '2026-11',
        departmentId: 'dept-uuid-1',
      );

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules/export-excel'));
      expect(fakeApiClient.capturedQueryParams?['target_month'], equals('2026-11'));
      expect(result, equals(sampleBytes));
    });
  });
}
