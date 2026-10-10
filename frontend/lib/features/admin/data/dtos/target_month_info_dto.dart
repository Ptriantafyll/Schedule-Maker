import 'package:frontend/features/admin/domain/models/target_month_info.dart';

class TargetMonthInfoDto {
  const TargetMonthInfoDto({
    required this.nextTargetMonth,
    this.lastPublishedMonth,
  });

  final String nextTargetMonth;
  final String? lastPublishedMonth;

  factory TargetMonthInfoDto.fromJson(Map<String, dynamic> json) {
    final nextTargetMonth = json['next_target_month'];
    if (nextTargetMonth == null || nextTargetMonth is! String) {
      throw const FormatException(
        'Missing required "next_target_month" in target month JSON',
      );
    }

    final lastPublishedMonth = json['last_published_month'] as String?;

    return TargetMonthInfoDto(
      nextTargetMonth: nextTargetMonth,
      lastPublishedMonth: lastPublishedMonth,
    );
  }

  TargetMonthInfo toDomain() {
    return TargetMonthInfo(
      nextTargetMonth: nextTargetMonth,
      lastPublishedMonth: lastPublishedMonth,
    );
  }

  Map<String, dynamic> toJson() {
    return {
      'next_target_month': nextTargetMonth,
      'last_published_month': lastPublishedMonth,
    };
  }
}
