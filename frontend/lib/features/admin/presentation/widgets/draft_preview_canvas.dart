import 'package:flutter/material.dart';
import 'package:frontend/features/admin/domain/models/schedule_draft.dart';
import 'package:skeletonizer/skeletonizer.dart';

enum CanvasViewMode { calendar, table }

class DraftPreviewCanvas extends StatefulWidget {
  const DraftPreviewCanvas({
    super.key,
    this.isGenerated = false,
    this.draft,
    this.isExporting = false,
    this.onExportPressed,
    this.generatedContent,
    this.isPublishing = false,
    this.onPublishPressed,
    this.isLoading = false,
  });

  final bool isGenerated;
  final ScheduleDraft? draft;
  final bool isExporting;
  final VoidCallback? onExportPressed;
  final Widget? generatedContent;
  final bool isPublishing;
  final VoidCallback? onPublishPressed;
  final bool isLoading;

  @override
  State<DraftPreviewCanvas> createState() => _DraftPreviewCanvasState();
}

class _DraftPreviewCanvasState extends State<DraftPreviewCanvas> {
  CanvasViewMode _viewMode = CanvasViewMode.calendar;
  int _selectedDay = 1;
  static final _skeletonDraft = ScheduleDraft(
    id: 'skeleton',
    targetMonth: '2026-11',
    status: 'draft',
    sourceFilename: 'Schedule Template.xlsx',
    totalDuties: 45,
    solverStatus: 'OPTIMAL',
    assignments: List.generate(
      15,
      (i) => ScheduleAssignment(
        date: '2026-11-${(i + 1).toString().padLeft(2, '0')}',
        dayName: 'Mon',
        doctorName: 'Dr. Gregory House',
        doctorEmail: 'gregory.house@hospital.org',
        position: 'Emergency',
        shift: 'Day',
      ),
    ),
    unavailabilities: const {},
    createdAt: DateTime(2026, 11, 1),
    updatedAt: DateTime(2026, 11, 1),
    departmentId: '1234',
  );

  Widget _buildPublishButton({bool isCompact = false}) {
    final isAlreadyPublished = widget.draft?.isPublished ?? false;
    final isEnabled =
        !widget.isLoading && !widget.isPublishing && !isAlreadyPublished;
    final onPressed = isEnabled ? widget.onPublishPressed : null;

    if (isCompact) {
      if (widget.isPublishing) {
        return const SizedBox(
          width: 24,
          height: 24,
          child: CircularProgressIndicator(strokeWidth: 2),
        );
      }
      return IconButton(
        onPressed: onPressed,
        tooltip: 'Publish Schedule',
        icon: const Icon(Icons.cloud_upload_outlined),
      );
    }

    return FilledButton.icon(
      icon: widget.isPublishing
          ? const SizedBox(
              width: 14,
              height: 14,
              child: CircularProgressIndicator(
                strokeWidth: 2,
                color: Colors.white,
              ),
            )
          : const Icon(Icons.cloud_upload_outlined, size: 18),
      onPressed: onPressed,
      label: const Text('Publish Schedule'),
    );
  }

  Widget _buildExportButton({bool isCompact = false}) {
    if (isCompact) {
      if (widget.isExporting) {
        return const SizedBox(
          width: 24,
          height: 24,
          child: CircularProgressIndicator(strokeWidth: 2),
        );
      }
      return IconButton(
        onPressed: widget.isLoading ? null : widget.onExportPressed,
        tooltip: 'Export to Excel',
        icon: const Icon(Icons.file_download_outlined),
      );
    }

    return OutlinedButton.icon(
      onPressed:
          (widget.isLoading || widget.isExporting) ? null : widget.onExportPressed,
      label: const Text('Export to Excel'),
      icon: widget.isExporting
          ? const SizedBox(
              width: 14,
              height: 14,
              child: CircularProgressIndicator(strokeWidth: 2),
            )
          : const Icon(Icons.file_download_outlined, size: 18),
    );
  }

  Widget _buildSummaryBar(ThemeData theme, ScheduleDraft draft) {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
      child: Wrap(
        spacing: 12,
        runSpacing: 8,
        crossAxisAlignment: WrapCrossAlignment.center,
        children: [
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
            decoration: BoxDecoration(
              color: widget.isLoading
                  ? theme.colorScheme.surfaceContainerHighest
                  : Colors.green.withValues(alpha: 0.15),
              borderRadius: BorderRadius.circular(16),
              border: Border.all(
                color: widget.isLoading
                    ? theme.colorScheme.outlineVariant
                    : Colors.green.shade600,
              ),
            ),
            child: Text(
              draft.solverStatus,
              style: theme.textTheme.labelMedium?.copyWith(
                color: widget.isLoading
                    ? theme.colorScheme.onSurfaceVariant
                    : Colors.green.shade800,
                fontWeight: FontWeight.bold,
              ),
            ),
          ),
          Text(
            '${draft.totalDuties} Duties',
            style: theme.textTheme.bodyMedium?.copyWith(
              fontWeight: FontWeight.w600,
            ),
          ),
          if (draft.sourceFilename.isNotEmpty)
            Text(
              draft.sourceFilename,
              style: theme.textTheme.bodySmall?.copyWith(
                color: theme.colorScheme.onSurfaceVariant,
              ),
            ),
        ],
      ),
    );
  }

  Widget _buildTable(ThemeData theme, ScheduleDraft draft) {
    if (draft.assignments.isEmpty) {
      return const Padding(
        padding: EdgeInsets.all(24),
        child: Center(child: Text('No shift assignments found')),
      );
    }

    return SingleChildScrollView(
      scrollDirection: Axis.horizontal,
      child: DataTable(
        columns: const [
          DataColumn(
            label: Text('Date', style: TextStyle(fontWeight: FontWeight.bold)),
          ),
          DataColumn(
            label: Text('Day', style: TextStyle(fontWeight: FontWeight.bold)),
          ),
          DataColumn(
            label: Text(
              'Doctor',
              style: TextStyle(fontWeight: FontWeight.bold),
            ),
          ),
          DataColumn(
            label: Text(
              'Position',
              style: TextStyle(fontWeight: FontWeight.bold),
            ),
          ),
          DataColumn(
            label: Text('Shift', style: TextStyle(fontWeight: FontWeight.bold)),
          ),
        ],
        rows: draft.assignments.map((assignment) {
          return DataRow(
            cells: [
              DataCell(Text(assignment.date)),
              DataCell(Text(assignment.dayName)),
              DataCell(Text(assignment.doctorName)),
              DataCell(Text(assignment.position)),
              DataCell(Text(assignment.shift)),
            ],
          );
        }).toList(),
      ),
    );
  }

  Widget _buildWeekdayHeader(ThemeData theme) {
    const weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

    return Padding(
      padding: const EdgeInsets.only(bottom: 8),
      child: Row(
        children: weekdays.map((day) {
          return Expanded(
            child: Center(
              child: Text(
                day,
                style: theme.textTheme.labelMedium?.copyWith(
                  fontWeight: FontWeight.bold,
                  color: theme.colorScheme.onSurfaceVariant,
                ),
              ),
            ),
          );
        }).toList(),
      ),
    );
  }

  Widget _buildShiftDots(
    ThemeData theme,
    String dateString,
    List<ScheduleAssignment> assignments,
  ) {
    if (assignments.isEmpty) {
      return const SizedBox(height: 8);
    }

    final dotColor = widget.isLoading
        ? theme.colorScheme.outlineVariant.withValues(alpha: 0.5)
        : Colors.green;

    return Wrap(
      alignment: WrapAlignment.center,
      spacing: 2,
      runSpacing: 2,
      children: assignments.asMap().entries.map((entry) {
        return Container(
          key: ValueKey('shift_dot_${dateString}_${entry.key}'),
          width: 6,
          height: 6,
          decoration: BoxDecoration(
            color: dotColor,
            shape: BoxShape.circle,
          ),
        );
      }).toList(),
    );
  }

  Widget _buildDayCell(
    ThemeData theme,
    int year,
    int month,
    int day,
    List<ScheduleAssignment> dayAssignments,
  ) {
    final dateString =
        '$year-${month.toString().padLeft(2, '0')}-${day.toString().padLeft(2, '0')}';
    final isSelected = !widget.isLoading && _selectedDay == day;

    return InkWell(
      key: ValueKey('cal_day_$day'),
      borderRadius: BorderRadius.circular(8),
      onTap: widget.isLoading ? null : () => setState(() => _selectedDay = day),
      child: Container(
        decoration: BoxDecoration(
          borderRadius: BorderRadius.circular(8),
          color: isSelected
              ? theme.colorScheme.primaryContainer.withValues(alpha: 0.5)
              : null,
          border: Border.all(
            color: isSelected
                ? theme.colorScheme.primary
                : theme.colorScheme.outlineVariant.withValues(alpha: 0.4),
          ),
        ),
        padding: const EdgeInsets.symmetric(vertical: 4, horizontal: 2),
        child: Column(
          mainAxisAlignment: MainAxisAlignment.spaceBetween,
          children: [
            Text(
              '$day',
              style: theme.textTheme.bodyMedium?.copyWith(
                fontWeight: isSelected ? FontWeight.bold : FontWeight.normal,
                color: isSelected ? theme.colorScheme.primary : null,
              ),
            ),
            _buildShiftDots(theme, dateString, dayAssignments),
          ],
        ),
      ),
    );
  }

  Widget _buildSelectedDayDetails(
    ThemeData theme,
    String selectedDateString,
    List<ScheduleAssignment> assignments,
  ) {
    return Container(
      width: double.infinity,
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        color: theme.colorScheme.surfaceContainerLow,
        borderRadius: BorderRadius.circular(8),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(
            'Assignments for $selectedDateString',
            style: theme.textTheme.titleSmall?.copyWith(
              fontWeight: FontWeight.bold,
            ),
          ),
          const SizedBox(height: 8),
          if (assignments.isEmpty)
            const Text('No shifts assigned for this date')
          else
            ...assignments.map((assignment) {
              return Card(
                elevation: 0,
                color: theme.colorScheme.surface,
                margin: const EdgeInsets.only(bottom: 6),
                child: ListTile(
                  dense: true,
                  leading: const CircleAvatar(
                    radius: 14,
                    child: Icon(Icons.person, size: 16),
                  ),
                  title: Text(
                    assignment.doctorName,
                    style: const TextStyle(fontWeight: FontWeight.bold),
                  ),
                  subtitle: Text(
                    '${assignment.position} • ${assignment.shift} Shift',
                  ),
                ),
              );
            }),
        ],
      ),
    );
  }

  Widget _buildGridItem(
    ThemeData theme,
    int year,
    int month,
    int index,
    int leadingBlanks,
    Map<String, List<ScheduleAssignment>> assignmentsByDate,
  ) {
    if (index < leadingBlanks) {
      return const SizedBox.shrink();
    }

    final day = index - leadingBlanks + 1;
    final dateString =
        '$year-${month.toString().padLeft(2, '0')}-${day.toString().padLeft(2, '0')}';
    final dayAssignments = assignmentsByDate[dateString] ?? const [];

    return _buildDayCell(theme, year, month, day, dayAssignments);
  }

  Widget _buildCalendarView(ThemeData theme, ScheduleDraft draft) {
    final parts = draft.targetMonth.split('-');
    final year = int.tryParse(parts[0]) ?? 2026;
    final month = int.tryParse(parts.length > 1 ? parts[1] : '11') ?? 11;

    final daysInMonth = DateTime(year, month + 1, 0).day;
    final firstWeekday = DateTime(year, month, 1).weekday;
    final leadingBlanks = firstWeekday - 1;

    final Map<String, List<ScheduleAssignment>> assignmentsByDate = {};
    for (final a in draft.assignments) {
      assignmentsByDate.putIfAbsent(a.date, () => []).add(a);
    }

    final selectedDateString =
        '$year-${month.toString().padLeft(2, '0')}-${_selectedDay.toString().padLeft(2, '0')}';
    final selectedDayAssignments =
        assignmentsByDate[selectedDateString] ?? const [];

    return Padding(
      padding: const EdgeInsets.all(16),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.stretch,
        children: [
          _buildWeekdayHeader(theme),
          GridView.builder(
            shrinkWrap: true,
            physics: const NeverScrollableScrollPhysics(),
            gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(
              crossAxisCount: 7,
              crossAxisSpacing: 4,
              mainAxisSpacing: 4,
              mainAxisExtent: 48,
            ),
            itemCount: leadingBlanks + daysInMonth,
            itemBuilder: (context, index) {
              return _buildGridItem(
                theme,
                year,
                month,
                index,
                leadingBlanks,
                assignmentsByDate,
              );
            },
          ),
          const SizedBox(height: 16),
          _buildSelectedDayDetails(
            theme,
            selectedDateString,
            selectedDayAssignments,
          ),
        ],
      ),
    );
  }

  Widget _buildEmptyState(BuildContext context) {
    final theme = Theme.of(context);

    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 56),
      child: Center(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            CircleAvatar(
              radius: 36,
              backgroundColor: theme.colorScheme.surfaceContainerHighest,
              child: Icon(
                Icons.calendar_month_outlined,
                size: 40,
                color: theme.colorScheme.onSurfaceVariant,
              ),
            ),
            const SizedBox(height: 16),
            Text(
              'Awaiting Generation',
              style: theme.textTheme.titleMedium?.copyWith(
                fontWeight: FontWeight.bold,
              ),
            ),
            const SizedBox(height: 8),
            ConstrainedBox(
              constraints: const BoxConstraints(maxWidth: 420),
              child: Text(
                'Click "Generate Schedule" above or import an Excel spreadsheet to produce optimized monthly assignments.',
                textAlign: TextAlign.center,
                style: theme.textTheme.bodyMedium?.copyWith(
                  color: theme.colorScheme.onSurfaceVariant,
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final isSkeleton = widget.isLoading && widget.draft == null;
    final effectiveDraft = isSkeleton ? _skeletonDraft : widget.draft;
    final hasDraft = widget.isGenerated || effectiveDraft != null;

    return Skeletonizer(
      enabled: widget.isLoading,
      child: SizedBox(
        width: double.infinity,
        child: Card(
          elevation: 0,
          shape: RoundedRectangleBorder(
            borderRadius: BorderRadius.circular(12),
            side: BorderSide(
              color: theme.colorScheme.outlineVariant.withValues(alpha: 0.6),
            ),
          ),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Padding(
                padding: const EdgeInsets.symmetric(
                  horizontal: 16,
                  vertical: 8,
                ),
                child: LayoutBuilder(
                  builder: (context, constraints) {
                    final isCompact = constraints.maxWidth < 500;
                    return Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        Flexible(
                          child: Text(
                            'Draft Preview Canvas',
                            overflow: TextOverflow.ellipsis,
                            style: theme.textTheme.titleMedium?.copyWith(
                              fontWeight: FontWeight.bold,
                            ),
                          ),
                        ),
                        const SizedBox(width: 8),
                        Row(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            if (hasDraft) ...[
                              _buildExportButton(isCompact: isCompact),
                              const SizedBox(width: 4),
                              _buildPublishButton(isCompact: isCompact),
                              const SizedBox(width: 4),
                            ],
                            IconButton(
                              isSelected: _viewMode == CanvasViewMode.calendar,
                              onPressed: widget.isLoading
                                  ? null
                                  : () => setState(
                                        () => _viewMode =
                                            CanvasViewMode.calendar,
                                      ),
                              icon: const Icon(Icons.calendar_view_month),
                              tooltip: 'Calendar View',
                            ),
                            IconButton(
                              isSelected: _viewMode == CanvasViewMode.table,
                              onPressed: widget.isLoading
                                  ? null
                                  : () => setState(
                                        () => _viewMode =
                                            CanvasViewMode.table,
                                      ),
                              icon: const Icon(Icons.table_chart_outlined),
                              tooltip: 'Table View',
                            ),
                          ],
                        ),
                      ],
                    );
                  },
                ),
              ),
              const Divider(height: 1),
              if (!hasDraft)
                _buildEmptyState(context)
              else if (effectiveDraft != null)
                Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    _buildSummaryBar(theme, effectiveDraft),
                    const Divider(height: 1),
                    _viewMode == CanvasViewMode.table
                        ? _buildTable(theme, effectiveDraft)
                        : _buildCalendarView(theme, effectiveDraft),
                  ],
                )
              else
                Padding(
                  padding: const EdgeInsets.all(16),
                  child:
                      widget.generatedContent ??
                      const Center(child: Text('Draft Schedule Generated')),
                ),
            ],
          ),
        ),
      ),
    );
  }
}
