import 'package:dio/dio.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/data/datasources/roster_remote_data_source.dart';
import 'package:frontend/features/auth/domain/models/user_role.dart';

class FakeApiClient implements ApiClient {
  String? capturedMethod;
  String? capturedPath;
  Object? capturedData;
  Map<String, dynamic>? capturedQueryParams;
  String? capturedContentType;
  bool? capturedRequiresAuth;

  Response<dynamic>? responseToReturn;
  Exception? exceptionToThrow;

  Future<Response<T>> _record<T>(
    String method,
    String path, {
    Object? data,
    Map<String, dynamic>? queryParameters,
    String? contentType,
    bool requiresAuth = true,
  }) async {
    capturedMethod = method;
    capturedPath = path;
    capturedData = data;
    capturedQueryParams = queryParameters;
    capturedContentType = contentType;
    capturedRequiresAuth = requiresAuth;

    if (exceptionToThrow != null) throw exceptionToThrow!;
    final r = responseToReturn!;
    return Response<T>(
      requestOptions: r.requestOptions,
      statusCode: r.statusCode,
      data: r.data as T?,
    );
  }

  @override
  Future<Response<T>> get<T>(
    String path, {
    Map<String, dynamic>? queryParameters,
    bool requiresAuth = true,
    ResponseType? responseType,
  }) => _record(
    'GET',
    path,
    queryParameters: queryParameters,
    requiresAuth: requiresAuth,
  );

  @override
  Future<Response<T>> post<T>(
    String path, {
    Object? data,
    Map<String, dynamic>? queryParameters,
    String? contentType,
    bool requiresAuth = true,
  }) => _record(
    'POST',
    path,
    data: data,
    queryParameters: queryParameters,
    contentType: contentType,
    requiresAuth: requiresAuth,
  );

  @override
  Future<Response<T>> patch<T>(
    String path, {
    Object? data,
    Map<String, dynamic>? queryParameters,
    String? contentType,
    bool requiresAuth = true,
  }) => _record(
    'PATCH',
    path,
    data: data,
    queryParameters: queryParameters,
    contentType: contentType,
    requiresAuth: requiresAuth,
  );

  @override
  Future<Response<T>> delete<T>(
    String path, {
    Object? data,
    Map<String, dynamic>? queryParameters,
    String? contentType,
    bool requiresAuth = true,
  }) => _record(
    'DELETE',
    path,
    data: data,
    queryParameters: queryParameters,
    contentType: contentType,
    requiresAuth: requiresAuth,
  );
}

void main() {
  late FakeApiClient fakeApiClient;
  late RosterRemoteDataSource dataSource;

  final sampleTeamJson = <String, dynamic>{
    'id': 'team-1',
    'name': 'Team A',
    'department_id': 'dept-1',
  };

  final samplePositionJson = <String, dynamic>{
    'id': 'pos-1',
    'name': 'ER',
    'department_id': 'dept-1',
    'duty_days': [0, 2, 4],
  };

  final sampleShiftJson = <String, dynamic>{
    'id': 'shift-1',
    'name': 'Day',
    'position_id': 'pos-1',
    'doctors_per_shift': 2,
    'grants_day_off': true,
  };

  final sampleDoctorJson = <String, dynamic>{
    'id': 'doc-1',
    'name': 'Dr. Panos',
    'department_id': 'dept-1',
    'team_id': 'team-1',
    'position_ids': ['pos-1'],
  };

  final sampleUnavailJson = <String, dynamic>{
    'id': 'unavail-1',
    'doctor_id': 'doc-1',
    'date': '2026-11-12',
  };

  final samplePreAssignJson = <String, dynamic>{
    'id': 'pa-1',
    'doctor_id': 'doc-1',
    'shift_id': 'shift-1',
    'date': '2026-11-18',
  };

  setUp(() {
    fakeApiClient = FakeApiClient();
    dataSource = RosterRemoteDataSource(fakeApiClient);
  });

  group('Teams endpoints', () {
    test('fetchTeams GETs /api/v1/teams/ and parses domain list', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/teams/'),
        statusCode: 200,
        data: [sampleTeamJson],
      );

      final result = await dataSource.fetchTeams();

      expect(fakeApiClient.capturedMethod, 'GET');
      expect(fakeApiClient.capturedPath, '/api/v1/teams/');
      expect(result.single.name, 'Team A');
    });

    test('createTeam POSTs /api/v1/teams/ with {name}', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/teams/'),
        statusCode: 201,
        data: sampleTeamJson,
      );

      final result = await dataSource.createTeam(name: 'Team A');

      expect(fakeApiClient.capturedMethod, 'POST');
      expect(fakeApiClient.capturedPath, '/api/v1/teams/');
      expect(fakeApiClient.capturedData, {'name': 'Team A'});
      expect(result.id, 'team-1');
    });

    test('updateTeam PATCHes /api/v1/teams/{teamId} with {name}', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/teams/team-1'),
        statusCode: 200,
        data: sampleTeamJson,
      );

      final result = await dataSource.updateTeam(
        teamId: 'team-1',
        name: 'Team A',
      );

      expect(fakeApiClient.capturedMethod, 'PATCH');
      expect(fakeApiClient.capturedPath, '/api/v1/teams/team-1');
      expect(fakeApiClient.capturedData, {'name': 'Team A'});
      expect(result.name, 'Team A');
    });

    test('deleteTeam DELETEs /api/v1/teams/{teamId}', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/teams/team-1'),
        statusCode: 204,
      );

      await dataSource.deleteTeam(teamId: 'team-1');

      expect(fakeApiClient.capturedMethod, 'DELETE');
      expect(fakeApiClient.capturedPath, '/api/v1/teams/team-1');
    });
  });

  group('Positions endpoints', () {
    test('fetchPositions GETs /api/v1/positions/', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/positions/'),
        statusCode: 200,
        data: [samplePositionJson],
      );

      final result = await dataSource.fetchPositions();

      expect(fakeApiClient.capturedMethod, 'GET');
      expect(fakeApiClient.capturedPath, '/api/v1/positions/');
      expect(result.single.dutyDays, [0, 2, 4]);
    });

    test('createPosition POSTs /api/v1/positions/ with snake_case body', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/positions/'),
        statusCode: 201,
        data: samplePositionJson,
      );

      final result = await dataSource.createPosition(
        name: 'ER',
        dutyDays: [0, 2, 4],
      );

      expect(fakeApiClient.capturedMethod, 'POST');
      expect(fakeApiClient.capturedPath, '/api/v1/positions/');
      expect(fakeApiClient.capturedData, {
        'name': 'ER',
        'duty_days': [0, 2, 4],
      });
      expect(result.id, 'pos-1');
    });

    test('updatePosition PATCHes only provided fields on /api/v1/positions/{id}', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/positions/pos-1'),
        statusCode: 200,
        data: samplePositionJson,
      );

      await dataSource.updatePosition(positionId: 'pos-1', dutyDays: [0, 2, 4]);

      expect(fakeApiClient.capturedMethod, 'PATCH');
      expect(fakeApiClient.capturedPath, '/api/v1/positions/pos-1');
      expect(fakeApiClient.capturedData, {
        'duty_days': [0, 2, 4],
      });
    });

    test('deletePosition DELETEs /api/v1/positions/{id}', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/positions/pos-1'),
        statusCode: 204,
      );

      await dataSource.deletePosition(positionId: 'pos-1');

      expect(fakeApiClient.capturedMethod, 'DELETE');
      expect(fakeApiClient.capturedPath, '/api/v1/positions/pos-1');
    });
  });

  group('Shifts endpoints', () {
    test('fetchShifts GETs /api/v1/shifts/', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/shifts/'),
        statusCode: 200,
        data: [sampleShiftJson],
      );

      final result = await dataSource.fetchShifts();

      expect(fakeApiClient.capturedMethod, 'GET');
      expect(fakeApiClient.capturedPath, '/api/v1/shifts/');
      expect(result.single.grantsDayOff, isTrue);
    });

    test('createShift POSTs /api/v1/shifts/ with position_id', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/shifts/'),
        statusCode: 201,
        data: sampleShiftJson,
      );

      await dataSource.createShift(
        positionId: 'pos-1',
        name: 'Day',
        doctorsPerShift: 2,
        grantsDayOff: true,
      );

      expect(fakeApiClient.capturedMethod, 'POST');
      expect(fakeApiClient.capturedPath, '/api/v1/shifts/');
      expect(fakeApiClient.capturedData, {
        'name': 'Day',
        'position_id': 'pos-1',
        'doctors_per_shift': 2,
        'grants_day_off': true,
      });
    });

    test('updateShift PATCHes /api/v1/shifts/{id} without position_id', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/shifts/shift-1'),
        statusCode: 200,
        data: sampleShiftJson,
      );

      await dataSource.updateShift(
        shiftId: 'shift-1',
        doctorsPerShift: 3,
        grantsDayOff: false,
      );

      expect(fakeApiClient.capturedMethod, 'PATCH');
      expect(fakeApiClient.capturedPath, '/api/v1/shifts/shift-1');
      expect(fakeApiClient.capturedData, {
        'doctors_per_shift': 3,
        'grants_day_off': false,
      });
    });

    test('deleteShift DELETEs /api/v1/shifts/{id}', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/shifts/shift-1'),
        statusCode: 204,
      );

      await dataSource.deleteShift(shiftId: 'shift-1');

      expect(fakeApiClient.capturedMethod, 'DELETE');
      expect(fakeApiClient.capturedPath, '/api/v1/shifts/shift-1');
    });
  });

  group('Doctors & sub-resource endpoints', () {
    test('fetchDoctors GETs /api/v1/doctors/ and parses position_ids', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/doctors/'),
        statusCode: 200,
        data: [sampleDoctorJson],
      );

      final result = await dataSource.fetchDoctors();

      expect(fakeApiClient.capturedMethod, 'GET');
      expect(fakeApiClient.capturedPath, '/api/v1/doctors/');
      expect(result.single.positionIds, ['pos-1']);
    });

    test('createDoctor POSTs /api/v1/doctors/ omitting null team_id', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/doctors/'),
        statusCode: 201,
        data: sampleDoctorJson,
      );

      await dataSource.createDoctor(name: 'Dr. Panos');

      expect(fakeApiClient.capturedMethod, 'POST');
      expect(fakeApiClient.capturedPath, '/api/v1/doctors/');
      expect(fakeApiClient.capturedData, {'name': 'Dr. Panos'});
    });

    test('updateDoctor PATCH omits team_id when unchanged and sends null when clearTeam=true', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/doctors/doc-1'),
        statusCode: 200,
        data: sampleDoctorJson,
      );

      await dataSource.updateDoctor(doctorId: 'doc-1', name: 'Dr. Renamed');
      expect(fakeApiClient.capturedMethod, 'PATCH');
      expect(fakeApiClient.capturedPath, '/api/v1/doctors/doc-1');
      expect(fakeApiClient.capturedData, {'name': 'Dr. Renamed'});

      await dataSource.updateDoctor(doctorId: 'doc-1', clearTeam: true);
      expect(fakeApiClient.capturedData, {'team_id': null});

      await dataSource.updateDoctor(doctorId: 'doc-1', teamId: 'team-2');
      expect(fakeApiClient.capturedData, {'team_id': 'team-2'});
    });

    test('deleteDoctor DELETEs /api/v1/doctors/{id}', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/doctors/doc-1'),
        statusCode: 204,
      );

      await dataSource.deleteDoctor(doctorId: 'doc-1');

      expect(fakeApiClient.capturedMethod, 'DELETE');
      expect(fakeApiClient.capturedPath, '/api/v1/doctors/doc-1');
    });

    test('assignPositionToDoctor POSTs /api/v1/doctors/{id}/position', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/doctors/doc-1/position'),
        statusCode: 201,
        data: {'id': 'link-1', 'doctor_id': 'doc-1', 'position_id': 'pos-1'},
      );

      await dataSource.assignPositionToDoctor(
        doctorId: 'doc-1',
        positionId: 'pos-1',
      );

      expect(fakeApiClient.capturedMethod, 'POST');
      expect(fakeApiClient.capturedPath, '/api/v1/doctors/doc-1/position');
      expect(fakeApiClient.capturedData, {'position_id': 'pos-1'});
    });

    test('removePositionFromDoctor DELETEs /api/v1/doctors/{id}/position/{posId}', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(
          path: '/api/v1/doctors/doc-1/position/pos-1',
        ),
        statusCode: 204,
      );

      await dataSource.removePositionFromDoctor(
        doctorId: 'doc-1',
        positionId: 'pos-1',
      );

      expect(fakeApiClient.capturedMethod, 'DELETE');
      expect(
        fakeApiClient.capturedPath,
        '/api/v1/doctors/doc-1/position/pos-1',
      );
    });

    test('unavailability fetch, add, and remove call exact /unavailability endpoints', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(
          path: '/api/v1/doctors/doc-1/unavailability',
        ),
        statusCode: 200,
        data: [sampleUnavailJson],
      );

      final list = await dataSource.fetchDoctorUnavailabilities(
        doctorId: 'doc-1',
      );
      expect(fakeApiClient.capturedMethod, 'GET');
      expect(
        fakeApiClient.capturedPath,
        '/api/v1/doctors/doc-1/unavailability',
      );
      expect(list.single.date, '2026-11-12');

      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(
          path: '/api/v1/doctors/doc-1/unavailability',
        ),
        statusCode: 201,
        data: sampleUnavailJson,
      );
      final created = await dataSource.addDoctorUnavailability(
        doctorId: 'doc-1',
        date: '2026-11-12',
      );
      expect(fakeApiClient.capturedMethod, 'POST');
      expect(fakeApiClient.capturedData, {'date': '2026-11-12'});
      expect(created.id, 'unavail-1');

      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(
          path: '/api/v1/doctors/doc-1/unavailability/unavail-1',
        ),
        statusCode: 204,
      );
      await dataSource.removeDoctorUnavailability(
        doctorId: 'doc-1',
        unavailabilityId: 'unavail-1',
      );
      expect(fakeApiClient.capturedMethod, 'DELETE');
      expect(
        fakeApiClient.capturedPath,
        '/api/v1/doctors/doc-1/unavailability/unavail-1',
      );
    });

    test('pre-assignment fetch, add, and remove call exact /pre-assignments endpoints', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(
          path: '/api/v1/doctors/doc-1/pre-assignments',
        ),
        statusCode: 200,
        data: [samplePreAssignJson],
      );

      final list = await dataSource.fetchDoctorPreAssignments(
        doctorId: 'doc-1',
      );
      expect(fakeApiClient.capturedMethod, 'GET');
      expect(
        fakeApiClient.capturedPath,
        '/api/v1/doctors/doc-1/pre-assignments',
      );
      expect(list.single.shiftId, 'shift-1');

      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(
          path: '/api/v1/doctors/doc-1/pre-assignments',
        ),
        statusCode: 201,
        data: samplePreAssignJson,
      );
      final created = await dataSource.addDoctorPreAssignment(
        doctorId: 'doc-1',
        shiftId: 'shift-1',
        date: '2026-11-18',
      );
      expect(fakeApiClient.capturedMethod, 'POST');
      expect(fakeApiClient.capturedData, {
        'shift_id': 'shift-1',
        'date': '2026-11-18',
      });
      expect(created.id, 'pa-1');

      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(
          path: '/api/v1/doctors/doc-1/pre-assignments/pa-1',
        ),
        statusCode: 204,
      );
      await dataSource.removeDoctorPreAssignment(
        doctorId: 'doc-1',
        preAssignmentId: 'pa-1',
      );
      expect(fakeApiClient.capturedMethod, 'DELETE');
      expect(
        fakeApiClient.capturedPath,
        '/api/v1/doctors/doc-1/pre-assignments/pa-1',
      );
    });

    test('rethrows ApiException on error response', () async {
      fakeApiClient.exceptionToThrow = const ApiException(
        type: ApiErrorType.validation,
        statusCode: 422,
        message: 'Doctor cannot be marked unavailable on a pre-assigned day',
      );

      expect(
        () => dataSource.addDoctorUnavailability(
          doctorId: 'doc-1',
          date: '2026-11-18',
        ),
        throwsA(isA<ApiException>()),
      );
    });
  });

  group('Invitations & doctor-user linking endpoints', () {
    final sampleInvitationJson = <String, dynamic>{
      'id': 'inv-1',
      'role': 'doctor',
      'department_id': 'dept-1',
      'doctor_id': 'doc-1',
      'created_by_user_id': 'admin-1',
      'expires_at': '2099-12-01T10:00:00Z',
      'used_at': null,
      'revoked_at': null,
      'created_at': '2026-11-01T08:00:00Z',
      'updated_at': '2026-11-01T08:00:00Z',
      'is_deleted': false,
      'sync_status': false,
      'max_uses': 1,
      'used_count': 0,
    };

    final sampleCreatedInvitationJson = <String, dynamic>{
      'id': 'inv-2',
      'role': 'doctor',
      'department_id': 'dept-1',
      'doctor_id': null,
      'expires_at': '2099-12-01T10:00:00Z',
      'raw_token': 'secret-token-xyz',
      'max_uses': 5,
      'used_count': 0,
    };

    final sampleDoctorUserJson = <String, dynamic>{
      'id': 'usr-1',
      'full_name': 'Dr. Panos',
      'email': 'panos@hospital.org',
      'doctor_id': 'doc-1',
    };

    test('fetchInvitations GETs /api/v1/auth/invitations', () async {
      fakeApiClient.responseToReturn = Response(
        requestOptions: RequestOptions(path: '/api/v1/auth/invitations'),
        statusCode: 200,
        data: [sampleInvitationJson],
      );

      final result = await dataSource.fetchInvitations();

      expect(fakeApiClient.capturedMethod, 'GET');
      expect(fakeApiClient.capturedPath, '/api/v1/auth/invitations');
      expect(result.single.id, 'inv-1');
      expect(result.single.role, UserRole.doctor);
      expect(result.single.doctorId, 'doc-1');
    });

    test(
      'createStaffInvitation POSTs /api/v1/auth/invitations/staff omitting null doctor_id',
      () async {
        fakeApiClient.responseToReturn = Response(
          requestOptions: RequestOptions(
            path: '/api/v1/auth/invitations/staff',
          ),
          statusCode: 201,
          data: sampleCreatedInvitationJson,
        );

        final bulk = await dataSource.createStaffInvitation(maxUses: 5);
        expect(fakeApiClient.capturedMethod, 'POST');
        expect(fakeApiClient.capturedPath, '/api/v1/auth/invitations/staff');
        expect(fakeApiClient.capturedData, {'role': 'doctor', 'max_uses': 5});
        expect(bulk.rawToken, 'secret-token-xyz');
        expect(bulk.maxUses, 5);

        await dataSource.createStaffInvitation(
          role: 'doctor',
          doctorId: 'doc-1',
          maxUses: 1,
        );
        expect(fakeApiClient.capturedData, {
          'role': 'doctor',
          'doctor_id': 'doc-1',
          'max_uses': 1,
        });
      },
    );

    test(
      'revokeInvitation POSTs /api/v1/auth/invitations/{id}/revoke',
      () async {
        fakeApiClient.responseToReturn = Response(
          requestOptions: RequestOptions(
            path: '/api/v1/auth/invitations/inv-1/revoke',
          ),
          statusCode: 200,
          data: {
            ...sampleInvitationJson,
            'revoked_at': '2026-11-10T10:00:00Z',
          },
        );

        final revoked = await dataSource.revokeInvitation(
          invitationId: 'inv-1',
        );

        expect(fakeApiClient.capturedMethod, 'POST');
        expect(
          fakeApiClient.capturedPath,
          '/api/v1/auth/invitations/inv-1/revoke',
        );
        expect(revoked.isActive, isFalse);
      },
    );

    test(
      'fetchDepartmentDoctorUsers GETs /api/v1/doctors/users',
      () async {
        fakeApiClient.responseToReturn = Response(
          requestOptions: RequestOptions(path: '/api/v1/doctors/users'),
          statusCode: 200,
          data: [sampleDoctorUserJson],
        );

        final users = await dataSource.fetchDepartmentDoctorUsers();

        expect(fakeApiClient.capturedMethod, 'GET');
        expect(fakeApiClient.capturedPath, '/api/v1/doctors/users');
        expect(users.single.email, 'panos@hospital.org');
        expect(users.single.doctorId, 'doc-1');
      },
    );

    test(
      'linkDoctorToUser POSTs /api/v1/doctors/{id}/link-user with {user_id}',
      () async {
        fakeApiClient.responseToReturn = Response(
          requestOptions: RequestOptions(
            path: '/api/v1/doctors/doc-1/link-user',
          ),
          statusCode: 200,
          data: {
            ...sampleDoctorJson,
            'linked_user_id': 'usr-1',
            'linked_user_email': 'panos@hospital.org',
          },
        );

        final linked = await dataSource.linkDoctorToUser(
          doctorId: 'doc-1',
          userId: 'usr-1',
        );

        expect(fakeApiClient.capturedMethod, 'POST');
        expect(
          fakeApiClient.capturedPath,
          '/api/v1/doctors/doc-1/link-user',
        );
        expect(fakeApiClient.capturedData, {'user_id': 'usr-1'});
        expect(linked.linkedUserId, 'usr-1');
        expect(linked.linkedUserEmail, 'panos@hospital.org');
      },
    );
  });
}
