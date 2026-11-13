import 'dart:convert';
import 'package:dio/dio.dart';
import 'package:flutter/widgets.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/data/dtos/excel_roster_diff_dto.dart';
import 'package:frontend/features/admin/data/dtos/schedule_draft_dto.dart';
import 'package:frontend/features/admin/domain/models/excel_roster_diff.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';
import 'package:frontend/features/admin/data/dtos/schedule_summary_dto.dart';
import 'package:frontend/features/admin/data/dtos/target_month_info_dto.dart';
import 'package:frontend/features/admin/domain/models/schedule_summary.dart';
import 'package:frontend/features/admin/domain/models/target_month_info.dart';

final scheduleDraftRemoteDataSourceProvider =
    Provider<ScheduleDraftRemoteDataSource>((ref) {
      final apiClient = ref.watch(apiClientProvider);
      return ScheduleDraftRemoteDataSource(apiClient);
    });

class ScheduleDraftRemoteDataSource {
  ScheduleDraftRemoteDataSource(this._apiClient);

  final ApiClient _apiClient;

  String _resolveCountryCode(String? countryCode) {
    if (countryCode != null && countryCode.isNotEmpty) {
      return countryCode;
    }

    try {
      final detectedCode =
          WidgetsBinding.instance.platformDispatcher.locale.countryCode;

      if (detectedCode != null && detectedCode.isNotEmpty) {
        return detectedCode;
      }
    } catch (_) {}

    return 'GR';
  }

  Future<ScheduleDraft> generateFromRoster({
    required String targetMonth,
    String? departmentId,
    String? countryCode,
  }) async {
    final payload = {
      'target_month': targetMonth,
      'department_id': ?departmentId,
      'country_code': _resolveCountryCode(countryCode),
    };

    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/schedules/generate-from-roster',
      data: payload,
    );

    return ScheduleDraftDto.fromJson(response.data!).toDomain();
  }

  Future<ScheduleDraft> generateFromExcel({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
    String? countryCode,
    bool syncRoster = false,
    List<String>? removeDoctorIds,
    List<String>? ignoredNewDoctors,
  }) async {
    // 1. Read bytes asynchronously using file_picker 13.x API
    final bytes = await file.readAsBytes();
    final multipartFile = MultipartFile.fromBytes(bytes, filename: file.name);

    final effectiveCountryCode = _resolveCountryCode(countryCode);

    // 2. Build FormData matching backend requirements
    final formDataMap = <String, dynamic>{
      'file': multipartFile,
      'target_month': targetMonth,
      'department_id': ?departmentId,
      'country_code': effectiveCountryCode,
      'sync_roster': syncRoster.toString(),
    };

    if (removeDoctorIds != null && removeDoctorIds.isNotEmpty) {
      formDataMap['remove_doctor_ids'] = jsonEncode(removeDoctorIds);
    }
    if (ignoredNewDoctors != null && ignoredNewDoctors.isNotEmpty) {
      formDataMap['ignored_new_doctors'] = jsonEncode(ignoredNewDoctors);
    }

    final formData = FormData.fromMap(formDataMap);

    // 3. Post to backend solver endpoint
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/schedules/generate-from-excel',
      data: formData,
      contentType: 'multipart/form-data',
    );

    // 4. Parse DTO and return domain model directly
    return ScheduleDraftDto.fromJson(response.data!).toDomain();
  }

  Future<ExcelRosterDiff> previewExcelRosterDiff({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
    String? countryCode,
  }) async {
    final bytes = await file.readAsBytes();
    final multipartFile = MultipartFile.fromBytes(bytes, filename: file.name);

    final formData = FormData.fromMap({
      'file': multipartFile,
      'target_month': targetMonth,
      'department_id': ?departmentId,
    });

    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/schedules/preview-excel-roster-diff',
      data: formData,
      contentType: 'multipart/form-data',
    );

    return ExcelRosterDiffDto.fromJson(response.data!).toDomain();
  }

  Future<ScheduleDraft?> getActiveDraft({
    required String targetMonth,
    String? departmentId,
  }) async {
    final queryParams = <String, dynamic>{
      'target_month': targetMonth,
      'department_id': ?departmentId,
    };

    try {
      final response = await _apiClient.get(
        '/api/v1/schedules/draft',
        queryParameters: queryParams,
      );

      return ScheduleDraftDto.fromJson(response.data!).toDomain();
    } on ApiException catch (e) {
      if (e.statusCode == 404) return null;
      rethrow;
    }
  }

  Future<List<int>> exportExcel({
    required String draftId,
    String? countryCode,
  }) async {
    final effectiveCountryCode = _resolveCountryCode(countryCode);

    final queryParams = <String, dynamic>{
      'draft_id': draftId,
      'country_code': effectiveCountryCode,
    };

    final response = await _apiClient.get<List<int>>(
      '/api/v1/schedules/export-excel',
      queryParameters: queryParams,
      responseType: ResponseType.bytes,
    );

    return response.data ?? const <int>[];
  }

  Future<List<ScheduleSummary>> fetchScheduleHistory({
    String? departmentId,
  }) async {
    final queryParams = <String, dynamic>{'department_id': ?departmentId};

    final response = await _apiClient.get<List<dynamic>>(
      '/api/v1/schedules',
      queryParameters: queryParams.isEmpty ? null : queryParams,
    );

    final rawList = response.data ?? const <dynamic>[];
    return rawList
        .map(
          (item) => ScheduleSummaryDto.fromJson(
            item as Map<String, dynamic>,
          ).toDomain(),
        )
        .toList();
  }

  Future<TargetMonthInfo> fetchTargetMonthInfo({String? departmentId}) async {
    final queryParams = <String, dynamic>{'department_id': ?departmentId};

    final response = await _apiClient.get<Map<String, dynamic>>(
      '/api/v1/schedules/target-month',
      queryParameters: queryParams.isEmpty ? null : queryParams,
    );

    return TargetMonthInfoDto.fromJson(response.data!).toDomain();
  }

  Future<ScheduleDraft> publishScheduleDraft({required String draftId}) async {
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/schedules/$draftId/publish',
    );

    return ScheduleDraftDto.fromJson(response.data!).toDomain();
  }
}
