import 'package:file_picker/file_picker.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/admin/data/datasources/schedule_draft_remote_data_source.dart';
import 'package:frontend/features/admin/domain/models/excel_roster_diff.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';
import 'package:frontend/features/admin/domain/models/schedule_summary.dart';
import 'package:frontend/features/admin/domain/models/target_month_info.dart';

final scheduleDraftRepositoryProvider = Provider<ScheduleDraftRepository>((
  ref,
) {
  final remoteDataSource = ref.watch(scheduleDraftRemoteDataSourceProvider);
  return ScheduleDraftRepositoryImpl(remoteDataSource: remoteDataSource);
});

abstract class ScheduleDraftRepository {
  Future<ScheduleDraft> generateFromExcel({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
    String? countryCode,
    bool syncRoster = false,
    List<String>? removeDoctorIds,
    List<String>? ignoredNewDoctors,
  });

  Future<ExcelRosterDiff> previewExcelRosterDiff({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
    String? countryCode,
  });

  Future<ScheduleDraft> generateFromRoster({
    required String targetMonth,
    String? departmentId,
    String? countryCode,
  });

  Future<ScheduleDraft?> getActiveDraft({
    required String targetMonth,
    String? departmentId,
  });

  Future<List<int>> exportExcel({required String draftId});

  Future<List<ScheduleSummary>> fetchScheduleHistory({String? departmentId});

  Future<TargetMonthInfo> fetchTargetMonthInfo({String? departmentId});

  Future<ScheduleDraft> publishScheduleDraft({required String draftId});
}

class ScheduleDraftRepositoryImpl implements ScheduleDraftRepository {
  ScheduleDraftRepositoryImpl({required this.remoteDataSource});

  final ScheduleDraftRemoteDataSource remoteDataSource;

  @override
  Future<ScheduleDraft> generateFromExcel({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
    String? countryCode,
    bool syncRoster = false,
    List<String>? removeDoctorIds,
    List<String>? ignoredNewDoctors,
  }) async {
    return await remoteDataSource.generateFromExcel(
      file: file,
      targetMonth: targetMonth,
      departmentId: departmentId,
      countryCode: countryCode,
      syncRoster: syncRoster,
      removeDoctorIds: removeDoctorIds,
      ignoredNewDoctors: ignoredNewDoctors,
    );
  }

  @override
  Future<ExcelRosterDiff> previewExcelRosterDiff({
    required PlatformFile file,
    required String targetMonth,
    String? departmentId,
    String? countryCode,
  }) async {
    return await remoteDataSource.previewExcelRosterDiff(
      file: file,
      targetMonth: targetMonth,
      departmentId: departmentId,
      countryCode: countryCode,
    );
  }

  @override
  Future<ScheduleDraft> generateFromRoster({
    required String targetMonth,
    String? departmentId,
    String? countryCode,
  }) async {
    return await remoteDataSource.generateFromRoster(
      targetMonth: targetMonth,
      departmentId: departmentId,
      countryCode: countryCode,
    );
  }

  @override
  Future<ScheduleDraft?> getActiveDraft({
    required String targetMonth,
    String? departmentId,
  }) async {
    return await remoteDataSource.getActiveDraft(
      targetMonth: targetMonth,
      departmentId: departmentId,
    );
  }

  @override
  Future<List<int>> exportExcel({required String draftId}) async {
    return await remoteDataSource.exportExcel(draftId: draftId);
  }

  @override
  Future<List<ScheduleSummary>> fetchScheduleHistory({
    String? departmentId,
  }) async {
    return await remoteDataSource.fetchScheduleHistory(
      departmentId: departmentId,
    );
  }

  @override
  Future<TargetMonthInfo> fetchTargetMonthInfo({String? departmentId}) async {
    return await remoteDataSource.fetchTargetMonthInfo(
      departmentId: departmentId,
    );
  }

  @override
  Future<ScheduleDraft> publishScheduleDraft({required String draftId}) async {
    return await remoteDataSource.publishScheduleDraft(draftId: draftId);
  }
}
