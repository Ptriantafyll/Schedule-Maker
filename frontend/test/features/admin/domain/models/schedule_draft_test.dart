import 'package:flutter_test/flutter_test.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';

void main() {
  group('ScheduleAssignment Domain Model Tests', () {
    const assignment1 = ScheduleAssignment(
      date: '2026-11-01',
      dayName: 'Sunday',
      doctorName: 'Dr. Gregory House',
      doctorEmail: 'house@hospital.org',
      position: 'ER',
      shift: 'Night',
    );

    const assignment2 = ScheduleAssignment(
      date: '2026-11-01',
      dayName: 'Sunday',
      doctorName: 'Dr. Gregory House',
      doctorEmail: 'house@hospital.org',
      position: 'ER',
      shift: 'Night',
    );

    const differentAssignment = ScheduleAssignment(
      date: '2026-11-01',
      dayName: 'Sunday',
      doctorName: 'Dr. Gregory House',
      doctorEmail: 'house@hospital.org',
      position: 'ER',
      shift: 'Morning',
    );

    test('supports value equality and identical hashCodes', () {
      expect(assignment1, equals(assignment2));
      expect(assignment1.hashCode, equals(assignment2.hashCode));
      expect(assignment1, isNot(equals(differentAssignment)));
    });

    test('toString includes doctor name, date, and shift', () {
      final str = assignment1.toString();
      expect(str, contains('Dr. Gregory House'));
      expect(str, contains('2026-11-01'));
      expect(str, contains('Night'));
    });
  });

  group('ScheduleDraft Domain Model Tests', () {
    final createdAt = DateTime.parse('2026-11-01T08:00:00.000Z');
    final updatedAt = DateTime.parse('2026-11-01T08:30:00.000Z');

    const assignment = ScheduleAssignment(
      date: '2026-11-01',
      dayName: 'Sunday',
      doctorName: 'Dr. Gregory House',
      doctorEmail: 'house@hospital.org',
      position: 'ER',
      shift: 'Night',
    );

    final draft1 = ScheduleDraft(
      id: 'draft-uuid-1',
      departmentId: 'dept-uuid-1',
      targetMonth: '2026-11',
      sourceFilename: 'november_roster.xlsx',
      totalDuties: 28,
      solverStatus: 'OPTIMAL',
      status: 'draft',
      assignments: const [assignment],
      unavailabilities: const {
        'Dr. Gregory House': [3, 14, 22],
      },
      createdAt: createdAt,
      updatedAt: updatedAt,
    );

    final draft2 = ScheduleDraft(
      id: 'draft-uuid-1',
      departmentId: 'dept-uuid-1',
      targetMonth: '2026-11',
      sourceFilename: 'november_roster.xlsx',
      totalDuties: 28,
      solverStatus: 'OPTIMAL',
      status: 'draft',
      assignments: const [assignment],
      unavailabilities: const {
        'Dr. Gregory House': [3, 14, 22],
      },
      createdAt: createdAt,
      updatedAt: updatedAt,
    );

    test('supports value equality and identical hashCodes', () {
      expect(draft1, equals(draft2));
      expect(draft1.hashCode, equals(draft2.hashCode));
    });

    test('isPublished getter evaluates status string accurately', () {
      expect(draft1.isPublished, isFalse);

      final publishedDraft = ScheduleDraft(
        id: 'draft-uuid-2',
        departmentId: 'dept-uuid-1',
        targetMonth: '2026-11',
        sourceFilename: 'november_roster.xlsx',
        totalDuties: 28,
        solverStatus: 'OPTIMAL',
        status: 'published',
        assignments: const [assignment],
        unavailabilities: const {},
        createdAt: createdAt,
        updatedAt: updatedAt,
      );

      expect(publishedDraft.isPublished, isTrue);
    });

    test('toString includes id, month, totalDuties, and status', () {
      final str = draft1.toString();
      expect(str, contains('draft-uuid-1'));
      expect(str, contains('2026-11'));
      expect(str, contains('28'));
      expect(str, contains('draft'));
    });
  });
}
