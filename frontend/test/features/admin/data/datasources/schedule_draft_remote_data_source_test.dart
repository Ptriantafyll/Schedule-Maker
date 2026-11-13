import 'dart:typed_data';
// ignore: depend_on_referenced_packages
import 'package:cross_file/cross_file.dart';
import 'package:dio/dio.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/data/datasources/schedule_draft_remote_data_source.dart';
import 'package:frontend/features/admin/domain/models/excel_roster_diff.dart';
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

  ResponseType? capturedResponseType;

  @override
  Future<Response<T>> get<T>(
    String path, {
    Map<String, dynamic>? queryParameters,
    bool requiresAuth = true,
    ResponseType? responseType,
  }) async {
    capturedPath = path;
    capturedQueryParams = queryParameters;
    capturedRequiresAuth = requiresAuth;
    capturedResponseType = responseType;

    if (exceptionToThrow != null) {
      throw exceptionToThrow!;
    }

    return responseToReturn as Response<T>;
  }
}

void main() {
  TestWidgetsFlutterBinding.ensureInitialized();
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

    test('includes country_code in FormData when generating schedule from excel', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/generate-from-excel'),
        statusCode: 201,
        data: sampleDraftPayload,
      );

      await dataSource.generateFromExcel(
        file: testFile,
        targetMonth: '2026-11',
        countryCode: 'GR',
      );

      final formData = fakeApiClient.capturedData as FormData;
      expect(
        formData.fields.any((f) => f.key == 'country_code' && f.value == 'GR'),
        isTrue,
      );
    });

    test('includes sync_roster, remove_doctor_ids, and ignored_new_doctors in FormData when syncing roster', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/generate-from-excel'),
        statusCode: 201,
        data: sampleDraftPayload,
      );

      await dataSource.generateFromExcel(
        file: testFile,
        targetMonth: '2026-11',
        departmentId: 'dept-uuid-1',
        syncRoster: true,
        removeDoctorIds: ['doc-id-1', 'doc-id-2'],
        ignoredNewDoctors: ['Dr. Ignored'],
      );

      final formData = fakeApiClient.capturedData as FormData;
      expect(
        formData.fields.any((f) => f.key == 'sync_roster' && f.value == 'true'),
        isTrue,
      );
      expect(
        formData.fields.any((f) => f.key == 'remove_doctor_ids' && f.value == '["doc-id-1","doc-id-2"]'),
        isTrue,
      );
      expect(
        formData.fields.any((f) => f.key == 'ignored_new_doctors' && f.value == '["Dr. Ignored"]'),
        isTrue,
      );
    });
  });

  group('ScheduleDraftRemoteDataSource.previewExcelRosterDiff', () {
    final sampleDiffPayload = <String, dynamic>{
      'has_existing_roster': true,
      'missing_from_excel': [
        {
          'id': 'doc-id-1',
          'name': 'Dr. Bob Kelso',
          'team_name': 'Alpha',
        }
      ],
      'new_in_excel': [
        {
          'name': 'Dr. Allison Cameron',
          'team_name': 'Beta',
          'positions': ['ER'],
        }
      ],
    };

    test('sends multipart POST to /api/v1/schedules/preview-excel-roster-diff and returns ExcelRosterDiff', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/preview-excel-roster-diff'),
        statusCode: 200,
        data: sampleDiffPayload,
      );

      final result = await dataSource.previewExcelRosterDiff(
        file: testFile,
        targetMonth: '2026-11',
        departmentId: 'dept-uuid-1',
      );

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules/preview-excel-roster-diff'));
      expect(fakeApiClient.capturedContentType, equals('multipart/form-data'));
      expect(fakeApiClient.capturedData, isA<FormData>());

      final formData = fakeApiClient.capturedData as FormData;
      expect(formData.fields.any((f) => f.key == 'target_month' && f.value == '2026-11'), isTrue);
      expect(formData.fields.any((f) => f.key == 'department_id' && f.value == 'dept-uuid-1'), isTrue);
      expect(formData.files.any((f) => f.key == 'file' && f.value.filename == 'nov_roster.xlsx'), isTrue);

      expect(result, isA<ExcelRosterDiff>());
      expect(result.hasExistingRoster, isTrue);
      expect(result.missingFromExcel.length, 1);
      expect(result.missingFromExcel.first.name, 'Dr. Bob Kelso');
      expect(result.newInExcel.length, 1);
      expect(result.newInExcel.first.name, 'Dr. Allison Cameron');
      expect(result.requiresReconciliation, isTrue);
    });

    test('propagates ApiException when diff preview fails', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.network,
        statusCode: 500,
        message: 'Diff error',
      );

      expect(
        () => dataSource.previewExcelRosterDiff(
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
        draftId: 'draft-uuid-1',
      );

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules/export-excel'));
      expect(fakeApiClient.capturedQueryParams?['draft_id'], equals('draft-uuid-1'));
      expect(fakeApiClient.capturedResponseType, equals(ResponseType.bytes));
      expect(result, equals(sampleBytes));
    });

    test('includes country_code in query parameters when exporting excel', () async {
      fakeApiClient.responseToReturn = Response<List<int>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/export-excel'),
        statusCode: 200,
        data: const [1, 2, 3],
      );

      await dataSource.exportExcel(
        draftId: 'draft-uuid-1',
        countryCode: 'GR',
      );

      expect(fakeApiClient.capturedQueryParams?['country_code'], equals('GR'));
    });
  });

  group('ScheduleDraftRemoteDataSource.fetchScheduleHistory', () {
    test('sends GET /api/v1/schedules with department_id and parses list of ScheduleSummary', () async {
      final sampleList = [
        {
          'id': 'summary-1',
          'department_id': 'dept-1',
          'target_month': '2026-11',
          'source_filename': 'nov.xlsx',
          'total_duties': 28,
          'solver_status': 'OPTIMAL',
          'status': 'published',
          'created_at': '2026-11-01T08:00:00.000Z',
          'updated_at': '2026-11-01T08:30:00.000Z',
        },
      ];

      fakeApiClient.responseToReturn = Response<List<dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules'),
        statusCode: 200,
        data: sampleList,
      );

      final result = await dataSource.fetchScheduleHistory(departmentId: 'dept-1');

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules'));
      expect(fakeApiClient.capturedQueryParams?['department_id'], equals('dept-1'));
      expect(result.length, equals(1));
      expect(result.first.id, equals('summary-1'));
      expect(result.first.isPublished, isTrue);
    });

    test('omits department_id query parameter when null', () async {
      fakeApiClient.responseToReturn = Response<List<dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules'),
        statusCode: 200,
        data: const [],
      );

      final result = await dataSource.fetchScheduleHistory();

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules'));
      expect(fakeApiClient.capturedQueryParams?['department_id'], isNull);
      expect(result, isEmpty);
    });

    test('propagates ApiException on server error', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.server,
        statusCode: 500,
        message: 'Internal server error',
      );

      expect(
        () => dataSource.fetchScheduleHistory(),
        throwsA(isA<ApiException>()),
      );
    });
  });

  group('ScheduleDraftRemoteDataSource.fetchTargetMonthInfo', () {
    test('sends GET /api/v1/schedules/target-month and parses TargetMonthInfo', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/target-month'),
        statusCode: 200,
        data: {
          'next_target_month': '2026-12',
          'last_published_month': '2026-11',
        },
      );

      final result = await dataSource.fetchTargetMonthInfo(departmentId: 'dept-1');

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules/target-month'));
      expect(fakeApiClient.capturedQueryParams?['department_id'], equals('dept-1'));
      expect(result.nextTargetMonth, equals('2026-12'));
      expect(result.lastPublishedMonth, equals('2026-11'));
      expect(result.hasPublishedSchedules, isTrue);
    });

    test('propagates ApiException on error', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.unauthorized,
        statusCode: 401,
        message: 'Unauthorized',
      );

      expect(
        () => dataSource.fetchTargetMonthInfo(),
        throwsA(isA<ApiException>()),
      );
    });
  });

  group('ScheduleDraftRemoteDataSource.publishScheduleDraft', () {
    test('sends POST /api/v1/schedules/{id}/publish and parses updated draft', () async {
      final publishedDraftPayload = Map<String, dynamic>.from(sampleDraftPayload)
        ..['status'] = 'published';

      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/draft-uuid-1/publish'),
        statusCode: 200,
        data: publishedDraftPayload,
      );

      final result = await dataSource.publishScheduleDraft(draftId: 'draft-uuid-1');

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules/draft-uuid-1/publish'));
      expect(result.status, equals('published'));
      expect(result.isPublished, isTrue);
    });

    test('propagates ApiException when draft is already published (HTTP 400)', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.validation,
        statusCode: 400,
        message: 'Schedule draft is already published',
      );

      expect(
        () => dataSource.publishScheduleDraft(draftId: 'draft-uuid-1'),
        throwsA(isA<ApiException>()),
      );
    });
  });

  group('ScheduleDraftRemoteDataSource.generateFromRoster', () {
    test('sends JSON POST to /api/v1/schedules/generate-from-roster and returns ScheduleDraft', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/generate-from-roster'),
        statusCode: 201,
        data: sampleDraftPayload,
      );

      final result = await dataSource.generateFromRoster(
        targetMonth: '2026-11',
        departmentId: 'dept-uuid-1',
        countryCode: 'GR',
      );

      expect(fakeApiClient.capturedPath, equals('/api/v1/schedules/generate-from-roster'));
      expect(fakeApiClient.capturedData, isA<Map<String, dynamic>>());

      final data = fakeApiClient.capturedData as Map<String, dynamic>;
      expect(data['target_month'], equals('2026-11'));
      expect(data['department_id'], equals('dept-uuid-1'));
      expect(data['country_code'], equals('GR'));

      expect(result, isA<ScheduleDraft>());
      expect(result.id, equals('draft-uuid-1'));
      expect(result.assignments.length, equals(1));
    });

    test('omits department_id from payload when departmentId is null', () async {
      fakeApiClient.responseToReturn = Response<Map<String, dynamic>>(
        requestOptions: RequestOptions(path: '/api/v1/schedules/generate-from-roster'),
        statusCode: 201,
        data: sampleDraftPayload,
      );

      await dataSource.generateFromRoster(
        targetMonth: '2026-11',
      );

      final data = fakeApiClient.capturedData as Map<String, dynamic>;
      expect(data['target_month'], equals('2026-11'));
      expect(data.containsKey('department_id'), isFalse);
    });

    test('propagates ApiException when solver endpoint fails', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.validation,
        statusCode: 400,
        message: "Position 'ER' has no eligible doctors assigned.",
      );

      expect(
        () => dataSource.generateFromRoster(
          targetMonth: '2026-11',
        ),
        throwsA(isA<ApiException>()),
      );
    });
  });
}
