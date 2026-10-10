import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/domain/models/schedule_summary.dart';

void main() {
  group('ScheduleSummary Domain Model Tests', () {
    final createdAt = DateTime.parse('2026-11-01T08:00:00.000Z');
    final updatedAt = DateTime.parse('2026-11-01T08:30:00.000Z');

    final summary1 = ScheduleSummary(
      id: 'summary-uuid-1',
      departmentId: 'dept-uuid-1',
      targetMonth: '2026-11',
      sourceFilename: 'november_roster.xlsx',
      totalDuties: 45,
      solverStatus: 'OPTIMAL',
      status: 'draft',
      createdAt: createdAt,
      updatedAt: updatedAt,
    );

    final summary2 = ScheduleSummary(
      id: 'summary-uuid-1',
      departmentId: 'dept-uuid-1',
      targetMonth: '2026-11',
      sourceFilename: 'november_roster.xlsx',
      totalDuties: 45,
      solverStatus: 'OPTIMAL',
      status: 'draft',
      createdAt: createdAt,
      updatedAt: updatedAt,
    );

    final differentSummary = ScheduleSummary(
      id: 'summary-uuid-2',
      departmentId: 'dept-uuid-1',
      targetMonth: '2026-12',
      sourceFilename: 'december_roster.xlsx',
      totalDuties: 30,
      solverStatus: 'OPTIMAL',
      status: 'published',
      createdAt: createdAt,
      updatedAt: updatedAt,
    );

    test('supports value equality and identical hashCodes', () {
      expect(summary1, equals(summary2));
      expect(summary1.hashCode, equals(summary2.hashCode));
      expect(summary1, isNot(equals(differentSummary)));
    });

    test('isPublished getter evaluates status string accurately', () {
      expect(summary1.isPublished, isFalse);
      expect(differentSummary.isPublished, isTrue);
    });

    test('toString includes key summary identifiers', () {
      final str = summary1.toString();
      expect(str, contains('summary-uuid-1'));
      expect(str, contains('2026-11'));
      expect(str, contains('draft'));
    });
  });
}
