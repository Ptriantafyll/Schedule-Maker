class TargetMonthInfo {
  const TargetMonthInfo({
    required this.nextTargetMonth,
    this.lastPublishedMonth,
  });

  final String nextTargetMonth;
  final String? lastPublishedMonth;

  bool get hasPublishedSchedules => lastPublishedMonth != null;

  @override
  bool operator ==(Object other) {
    if (identical(other, this)) return true;

    return other is TargetMonthInfo &&
        other.nextTargetMonth == nextTargetMonth &&
        other.lastPublishedMonth == lastPublishedMonth;
  }

  @override
  int get hashCode => Object.hash(nextTargetMonth, lastPublishedMonth);

  @override
  String toString() =>
      'TargetMonthInfo(nextTargetMonth: $nextTargetMonth, lastPublishedMonth: $lastPublishedMonth)';
}
