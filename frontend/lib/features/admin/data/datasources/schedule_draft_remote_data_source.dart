import 'package:dio/dio.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/core/network/api_client.dart';
import 'package:frontend/core/network/api_exception.dart';
import 'package:frontend/features/admin/data/dtos/schedule_draft_dto.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';

final scheduleDraftRemoteDataSourceProvider =
    Provider<ScheduleDraftRemoteDataSource>((ref) {
      final apiClient = ref.watch(apiClientProvider);
      return ScheduleDraftRemoteDataSource(apiClient);
    });

class ScheduleDraftRemoteDataSource {
  ScheduleDraftRemoteDataSource(this._apiClient);

  final ApiClient _apiClient;

  Future<ScheduleDraft> generateFromExcel({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
  }) async {
    // 1. Read bytes asynchronously using file_picker 13.x API
    final bytes = await file.readAsBytes();
    final multipartFile = MultipartFile.fromBytes(bytes, filename: file.name);

    // 2. Build FormData matching backend requirements
    final formData = FormData.fromMap({
      'file': multipartFile,
      'target_month': targetMonth,
      'department_id': ?departmentId,
    });

    // 3. Post to backend solver endpoint
    final response = await _apiClient.post<Map<String, dynamic>>(
      '/api/v1/schedules/generate-from-excel',
      data: formData,
      contentType: 'multipart/form-data',
    );

    // 4. Parse DTO and return domain model directly
    return ScheduleDraftDto.fromJson(response.data!).toDomain();
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
  }) async {
    final queryParams = <String, dynamic>{
      'draft_id': draftId,
    };

    final response = await _apiClient.get<List<int>>(
      '/api/v1/schedules/export-excel',
      queryParameters: queryParams,
      responseType: ResponseType.bytes,
    );

    return response.data ?? const <int>[];
  }
}
