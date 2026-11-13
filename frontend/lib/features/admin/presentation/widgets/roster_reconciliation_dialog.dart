import 'package:flutter/material.dart';
import 'package:frontend/features/admin/domain/models/excel_roster_diff.dart';

class RosterReconciliationChoices {
  const RosterReconciliationChoices({
    this.removeDoctorIds = const [],
    this.ignoredNewDoctors = const [],
  });

  final List<String> removeDoctorIds;
  final List<String> ignoredNewDoctors;

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;
    return other is RosterReconciliationChoices &&
        _listEquals(other.removeDoctorIds, removeDoctorIds) &&
        _listEquals(other.ignoredNewDoctors, ignoredNewDoctors);
  }

  @override
  int get hashCode => Object.hash(
        Object.hashAll(removeDoctorIds),
        Object.hashAll(ignoredNewDoctors),
      );
}

bool _listEquals<T>(List<T>? a, List<T>? b) {
  if (a == null) return b == null;
  if (b == null || a.length != b.length) return false;
  if (identical(a, b)) return true;
  for (int i = 0; i < a.length; i++) {
    if (a[i] != b[i]) return false;
  }
  return true;
}

class RosterReconciliationDialog extends StatefulWidget {
  const RosterReconciliationDialog({super.key, required this.diff});

  final ExcelRosterDiff diff;

  @override
  State<RosterReconciliationDialog> createState() =>
      _RosterReconciliationDialogState();
}

class _RosterReconciliationDialogState
    extends State<RosterReconciliationDialog> {
  final Map<String, bool> _keepDoctorMap = {};
  final Map<String, bool> _addDoctorMap = {};

  @override
  void initState() {
    super.initState();
    for (final doc in widget.diff.missingFromExcel) {
      _keepDoctorMap[doc.id] = true; // Default: Keep
    }
    for (final doc in widget.diff.newInExcel) {
      _addDoctorMap[doc.name] = true; // Default: Add
    }
  }

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);

    return AlertDialog(
      title: const Row(
        children: [
          Icon(Icons.published_with_changes_outlined),
          SizedBox(width: 8),
          Text('Reconcile Roster Changes'),
        ],
      ),
      content: ConstrainedBox(
        constraints: const BoxConstraints(
          maxWidth: 600,
          maxHeight: 420,
        ),
        child: SingleChildScrollView(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              Text(
                'We detected roster differences between the spreadsheet and your department database. Choose how to handle each doctor below before generating the schedule:',
                style: theme.textTheme.bodyMedium,
              ),
              const SizedBox(height: 16),
              if (widget.diff.missingFromExcel.isNotEmpty) ...[
                _buildMissingDoctorsSection(context),
                if (widget.diff.newInExcel.isNotEmpty)
                  const SizedBox(height: 16),
              ],
              if (widget.diff.newInExcel.isNotEmpty)
                _buildNewDoctorsSection(context),
            ],
          ),
        ),
      ),
      actions: [
        TextButton(
          onPressed: () => Navigator.of(context).pop(),
          child: const Text('Cancel'),
        ),
        FilledButton(
          onPressed: () {
            final removeDoctorIds = widget.diff.missingFromExcel
                .where((doc) => _keepDoctorMap[doc.id] == false)
                .map((doc) => doc.id)
                .toList();

            final ignoredNewDoctors = widget.diff.newInExcel
                .where((doc) => _addDoctorMap[doc.name] == false)
                .map((doc) => doc.name)
                .toList();

            Navigator.of(context).pop(
              RosterReconciliationChoices(
                removeDoctorIds: removeDoctorIds,
                ignoredNewDoctors: ignoredNewDoctors,
              ),
            );
          },
          child: const Text('Confirm & Generate'),
        ),
      ],
    );
  }

  Widget _buildMissingDoctorsSection(BuildContext context) {
    final theme = Theme.of(context);
    final count = widget.diff.missingFromExcel.length;

    return Card(
      elevation: 0,
      color: theme.colorScheme.surfaceContainerHighest.withValues(alpha: 0.4),
      margin: EdgeInsets.zero,
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            Row(
              children: [
                Icon(Icons.person_off_outlined, color: theme.colorScheme.error),
                const SizedBox(width: 8),
                Text(
                  'Missing from Spreadsheet ($count)',
                  style: theme.textTheme.titleMedium,
                ),
              ],
            ),
            const SizedBox(height: 4),
            Text(
              'These doctors exist in your department database but were not found in the spreadsheet. Choose whether to keep them in the roster or remove them.',
              style: theme.textTheme.bodySmall,
            ),
            const SizedBox(height: 12),
            ...widget.diff.missingFromExcel
                .map((doc) => _buildMissingDoctorRow(context, doc)),
          ],
        ),
      ),
    );
  }

  Widget _buildMissingDoctorRow(
    BuildContext context,
    MissingDoctorDiffItem doc,
  ) {
    final isKept = _keepDoctorMap[doc.id] ?? true;

    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 6),
      child: Row(
        children: [
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  doc.name,
                  style: const TextStyle(fontWeight: FontWeight.bold),
                ),
                Text(
                  doc.teamName != null ? 'Team ${doc.teamName}' : 'Teamless',
                  style: Theme.of(context).textTheme.bodySmall,
                ),
              ],
            ),
          ),
          SegmentedButton<bool>(
            segments: const [
              ButtonSegment<bool>(value: true, label: Text('Keep')),
              ButtonSegment<bool>(value: false, label: Text('Remove')),
            ],
            selected: {isKept},
            onSelectionChanged: (selected) {
              setState(() => _keepDoctorMap[doc.id] = selected.first);
            },
          ),
        ],
      ),
    );
  }

  Widget _buildNewDoctorsSection(BuildContext context) {
    final theme = Theme.of(context);
    final count = widget.diff.newInExcel.length;

    return Card(
      elevation: 0,
      color: theme.colorScheme.surfaceContainerHighest.withValues(alpha: 0.4),
      margin: EdgeInsets.zero,
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            Row(
              children: [
                Icon(
                  Icons.person_add_alt_outlined,
                  color: theme.colorScheme.primary,
                ),
                const SizedBox(width: 8),
                Text(
                  'New in Spreadsheet ($count)',
                  style: theme.textTheme.titleMedium,
                ),
              ],
            ),
            const SizedBox(height: 4),
            Text(
              'These doctors were found in the spreadsheet but are not yet in your department. Choose whether to add them to your roster or skip them.',
              style: theme.textTheme.bodySmall,
            ),
            const SizedBox(height: 12),
            ...widget.diff.newInExcel
                .map((doc) => _buildNewDoctorRow(context, doc)),
          ],
        ),
      ),
    );
  }

  Widget _buildNewDoctorRow(BuildContext context, NewDoctorDiffItem doc) {
    final isAdded = _addDoctorMap[doc.name] ?? true;

    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 6),
      child: Row(
        children: [
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  doc.name,
                  style: const TextStyle(fontWeight: FontWeight.bold),
                ),
                Text(
                  doc.teamName != null ? 'Team ${doc.teamName}' : 'Teamless',
                  style: Theme.of(context).textTheme.bodySmall,
                ),
                if (doc.positions.isNotEmpty)
                  Padding(
                    padding: const EdgeInsets.only(top: 4),
                    child: Wrap(
                      spacing: 4,
                      children: doc.positions
                          .map(
                            (p) => Chip(
                              label: Text(
                                p,
                                style: const TextStyle(fontSize: 11),
                              ),
                              visualDensity: VisualDensity.compact,
                              padding: EdgeInsets.zero,
                            ),
                          )
                          .toList(),
                    ),
                  ),
              ],
            ),
          ),
          SegmentedButton<bool>(
            segments: const [
              ButtonSegment<bool>(value: true, label: Text('Add')),
              ButtonSegment<bool>(value: false, label: Text('Skip')),
            ],
            selected: {isAdded},
            onSelectionChanged: (selected) {
              setState(() => _addDoctorMap[doc.name] = selected.first);
            },
          ),
        ],
      ),
    );
  }
}
