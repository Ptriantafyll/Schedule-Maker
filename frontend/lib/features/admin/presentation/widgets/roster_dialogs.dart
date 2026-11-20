import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';
import 'package:frontend/features/admin/presentation/controllers/roster_management_controller.dart';
import 'package:intl/intl.dart';

const List<String> kWeekdayLabels = [
  'Mon',
  'Tue',
  'Wed',
  'Thu',
  'Fri',
  'Sat',
  'Sun',
];

class _SelectableTile extends StatelessWidget {
  const _SelectableTile({
    required this.label,
    required this.selected,
    required this.onTap,
  });

  final String label;
  final bool selected;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    return Material(
      color: selected ? scheme.primaryContainer : Colors.transparent,
      borderRadius: BorderRadius.circular(10),
      child: InkWell(
        borderRadius: BorderRadius.circular(10),
        onTap: onTap,
        child: DecoratedBox(
          decoration: BoxDecoration(
            borderRadius: BorderRadius.circular(10),
            border: Border.all(
              color: selected ? scheme.primary : scheme.outlineVariant,
            ),
          ),
          child: Center(
            child: Text(
              label,
              style: Theme.of(context).textTheme.labelLarge?.copyWith(
                color: selected ? scheme.onPrimaryContainer : scheme.onSurface,
                fontWeight: selected ? FontWeight.w700 : FontWeight.w500,
              ),
            ),
          ),
        ),
      ),
    );
  }
}

class PositionFormResult {
  const PositionFormResult({required this.name, required this.dutyDays});
  final String name;
  final List<int> dutyDays;
}

class ShiftFormResult {
  const ShiftFormResult({
    required this.name,
    required this.doctorsPerShift,
    required this.grantsDayOff,
  });
  final String name;
  final int doctorsPerShift;
  final bool grantsDayOff;
}

class DoctorFormResult {
  const DoctorFormResult({
    required this.name,
    required this.teamId,
    required this.positionIds,
  });
  final String name;
  final String? teamId;
  final List<String> positionIds;
}

Future<bool> showRosterConfirmDeleteDialog(
  BuildContext context, {
  required String title,
  required String message,
}) async {
  final scheme = Theme.of(context).colorScheme;
  final result = await showDialog<bool>(
    context: context,
    builder: (dialogContext) => AlertDialog(
      title: Text(title),
      content: Text(message),
      actions: [
        TextButton(
          onPressed: () => Navigator.of(dialogContext).pop(false),
          child: const Text('Cancel'),
        ),
        FilledButton(
          style: FilledButton.styleFrom(
            backgroundColor: scheme.error,
            foregroundColor: scheme.onError,
          ),
          onPressed: () => Navigator.of(dialogContext).pop(true),
          child: const Text('Delete'),
        ),
      ],
    ),
  );
  return result ?? false;
}

Future<String?> showRosterTeamDialog(
  BuildContext context, {
  RosterTeam? initialTeam,
}) {
  final formKey = GlobalKey<FormState>();
  final controller = TextEditingController(text: initialTeam?.name ?? '');

  return showDialog<String>(
    context: context,
    builder: (dialogContext) => AlertDialog(
      title: Text(initialTeam == null ? 'Add Team' : 'Rename Team'),
      content: Form(
        key: formKey,
        child: TextFormField(
          controller: controller,
          autofocus: true,
          decoration: const InputDecoration(
            labelText: 'Team Name',
            hintText: 'e.g. Cardiology Alpha',
            border: OutlineInputBorder(),
          ),
          validator: (value) =>
              (value == null || value.trim().isEmpty) ? 'Required' : null,
        ),
      ),
      actions: [
        TextButton(
          onPressed: () => Navigator.of(dialogContext).pop(),
          child: const Text('Cancel'),
        ),
        FilledButton(
          onPressed: () {
            if (formKey.currentState!.validate()) {
              Navigator.of(dialogContext).pop(controller.text.trim());
            }
          },
          child: const Text('Save'),
        ),
      ],
    ),
  );
}

Future<PositionFormResult?> showRosterPositionDialog(
  BuildContext context, {
  RosterPosition? initialPosition,
}) {
  final formKey = GlobalKey<FormState>();
  final nameController = TextEditingController(
    text: initialPosition?.name ?? '',
  );
  final selectedDays = <int>{...(initialPosition?.dutyDays ?? const <int>[])};
  var showDaysError = false;

  return showDialog<PositionFormResult>(
    context: context,
    builder: (dialogContext) => StatefulBuilder(
      builder: (context, setState) {
        final scheme = Theme.of(context).colorScheme;
        return AlertDialog(
          title: Text(
            initialPosition == null ? 'Add Position' : 'Edit Position',
          ),
          content: Form(
            key: formKey,
            child: SingleChildScrollView(
              child: Column(
                mainAxisSize: MainAxisSize.min,
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  TextFormField(
                    controller: nameController,
                    autofocus: true,
                    decoration: const InputDecoration(
                      labelText: 'Position Name',
                      hintText: 'e.g. Emergency Room',
                      border: OutlineInputBorder(),
                    ),
                    validator: (v) =>
                        (v == null || v.trim().isEmpty) ? 'Required' : null,
                  ),
                  const SizedBox(height: 16),
                  Text(
                    'Active Duty Days',
                    style: Theme.of(context).textTheme.labelLarge,
                  ),
                  const SizedBox(height: 8),
                  for (var row = 0; row < 2; row++) ...[
                    if (row > 0) const SizedBox(height: 6),
                    Row(
                      children: [
                        for (var col = 0; col < 4; col++) ...[
                          if (col > 0) const SizedBox(width: 6),
                          Expanded(
                            child: (row * 4 + col) < kWeekdayLabels.length
                                ? SizedBox(
                                    height: 38,
                                    child: _SelectableTile(
                                      label: kWeekdayLabels[row * 4 + col],
                                      selected: selectedDays.contains(
                                        row * 4 + col,
                                      ),
                                      onTap: () {
                                        final day = row * 4 + col;
                                        setState(() {
                                          if (selectedDays.contains(day)) {
                                            selectedDays.remove(day);
                                          } else {
                                            selectedDays.add(day);
                                          }
                                          showDaysError = false;
                                        });
                                      },
                                    ),
                                  )
                                : const SizedBox(height: 38),
                          ),
                        ],
                      ],
                    ),
                  ],
                  if (showDaysError) ...[
                    const SizedBox(height: 8),
                    Text(
                      'Select at least one duty day.',
                      style: TextStyle(color: scheme.error, fontSize: 12),
                    ),
                  ],
                ],
              ),
            ),
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(dialogContext).pop(),
              child: const Text('Cancel'),
            ),
            FilledButton(
              onPressed: () {
                final validName = formKey.currentState!.validate();
                if (selectedDays.isEmpty) {
                  setState(() => showDaysError = true);
                  return;
                }
                if (validName) {
                  final sorted = selectedDays.toList()..sort();
                  Navigator.of(dialogContext).pop(
                    PositionFormResult(
                      name: nameController.text.trim(),
                      dutyDays: sorted,
                    ),
                  );
                }
              },
              child: const Text('Save'),
            ),
          ],
        );
      },
    ),
  );
}

Future<ShiftFormResult?> showRosterShiftDialog(
  BuildContext context, {
  RosterShift? initialShift,
}) {
  final formKey = GlobalKey<FormState>();
  final nameController = TextEditingController(text: initialShift?.name ?? '');
  final doctorsController = TextEditingController(
    text: (initialShift?.doctorsPerShift ?? 1).toString(),
  );
  var grantsDayOff = initialShift?.grantsDayOff ?? false;

  return showDialog<ShiftFormResult>(
    context: context,
    builder: (dialogContext) => StatefulBuilder(
      builder: (context, setState) => AlertDialog(
        title: Text(initialShift == null ? 'Add Shift' : 'Edit Shift'),
        content: Form(
          key: formKey,
          child: SingleChildScrollView(
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                TextFormField(
                  controller: nameController,
                  autofocus: true,
                  decoration: const InputDecoration(
                    labelText: 'Shift Name',
                    hintText: 'e.g. Night',
                    border: OutlineInputBorder(),
                  ),
                  validator: (v) =>
                      (v == null || v.trim().isEmpty) ? 'Required' : null,
                ),
                const SizedBox(height: 12),
                TextFormField(
                  controller: doctorsController,
                  keyboardType: TextInputType.number,
                  decoration: const InputDecoration(
                    labelText: 'Doctors Per Shift',
                    border: OutlineInputBorder(),
                  ),
                  validator: (v) {
                    final parsed = int.tryParse((v ?? '').trim());
                    if (parsed == null || parsed <= 0) {
                      return 'Enter a whole number >= 1';
                    }
                    return null;
                  },
                ),
                const SizedBox(height: 8),
                SwitchListTile(
                  contentPadding: EdgeInsets.zero,
                  title: const Text('Grants Day Off'),
                  subtitle: const Text('Doctor receives next day off'),
                  value: grantsDayOff,
                  onChanged: (v) => setState(() => grantsDayOff = v),
                ),
              ],
            ),
          ),
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(dialogContext).pop(),
            child: const Text('Cancel'),
          ),
          FilledButton(
            onPressed: () {
              if (formKey.currentState!.validate()) {
                Navigator.of(dialogContext).pop(
                  ShiftFormResult(
                    name: nameController.text.trim(),
                    doctorsPerShift: int.parse(doctorsController.text.trim()),
                    grantsDayOff: grantsDayOff,
                  ),
                );
              }
            },
            child: const Text('Save'),
          ),
        ],
      ),
    ),
  );
}

Future<DoctorFormResult?> showRosterDoctorDialog(
  BuildContext context, {
  RosterDoctor? initialDoctor,
  required List<RosterTeam> teams,
  required List<RosterPosition> positions,
}) {
  final formKey = GlobalKey<FormState>();
  final nameController = TextEditingController(text: initialDoctor?.name ?? '');
  String? selectedTeamId = initialDoctor?.teamId;
  final selectedPositions = <String>{
    ...(initialDoctor?.positionIds ?? const <String>[]),
  };

  return showDialog<DoctorFormResult>(
    context: context,
    builder: (dialogContext) => StatefulBuilder(
      builder: (context, setState) => AlertDialog(
        title: Text(initialDoctor == null ? 'Add Doctor' : 'Edit Doctor'),
        content: Form(
          key: formKey,
          child: SingleChildScrollView(
            child: Column(
              mainAxisSize: MainAxisSize.min,
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                TextFormField(
                  controller: nameController,
                  autofocus: true,
                  decoration: const InputDecoration(
                    labelText: 'Doctor Name',
                    hintText: 'e.g. Dr. Panos',
                    border: OutlineInputBorder(),
                  ),
                  validator: (v) =>
                      (v == null || v.trim().isEmpty) ? 'Required' : null,
                ),
                const SizedBox(height: 12),
                DropdownButtonFormField<String?>(
                  initialValue: teams.any((t) => t.id == selectedTeamId)
                      ? selectedTeamId
                      : null,
                  decoration: const InputDecoration(
                    labelText: 'Team',
                    border: OutlineInputBorder(),
                  ),
                  items: [
                    const DropdownMenuItem<String?>(
                      value: null,
                      child: Text('No Team'),
                    ),
                    for (final team in teams)
                      DropdownMenuItem<String?>(
                        value: team.id,
                        child: Text(team.name),
                      ),
                  ],
                  onChanged: (value) => setState(() => selectedTeamId = value),
                ),
                const SizedBox(height: 16),
                Text(
                  'Eligible Positions',
                  style: Theme.of(context).textTheme.labelLarge,
                ),
                const SizedBox(height: 8),
                if (positions.isEmpty)
                  Text(
                    'Create positions first to assign eligibility.',
                    style: Theme.of(context).textTheme.bodySmall,
                  )
                else
                  Wrap(
                    spacing: 6,
                    runSpacing: 6,
                    children: [
                      for (final pos in positions)
                        FilterChip(
                          label: Text(pos.name),
                          selected: selectedPositions.contains(pos.id),
                          onSelected: (selected) {
                            setState(() {
                              if (selected) {
                                selectedPositions.add(pos.id);
                              } else {
                                selectedPositions.remove(pos.id);
                              }
                            });
                          },
                        ),
                    ],
                  ),
              ],
            ),
          ),
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(dialogContext).pop(),
            child: const Text('Cancel'),
          ),
          FilledButton(
            onPressed: () {
              if (formKey.currentState!.validate()) {
                Navigator.of(dialogContext).pop(
                  DoctorFormResult(
                    name: nameController.text.trim(),
                    teamId: selectedTeamId,
                    positionIds: selectedPositions.toList(),
                  ),
                );
              }
            },
            child: const Text('Save'),
          ),
        ],
      ),
    ),
  );
}

Future<void> showDoctorDatesDialog(
  BuildContext context, {
  required RosterDoctor doctor,
  required List<RosterShift> shifts,
  required List<RosterPosition> positions,
}) {
  return showDialog<void>(
    context: context,
    builder: (_) => _DoctorDatesDialog(
      doctor: doctor,
      shifts: shifts,
      positions: positions,
    ),
  );
}

class _DoctorDatesDialog extends ConsumerStatefulWidget {
  const _DoctorDatesDialog({
    required this.doctor,
    required this.shifts,
    required this.positions,
  });

  final RosterDoctor doctor;
  final List<RosterShift> shifts;
  final List<RosterPosition> positions;

  @override
  ConsumerState<_DoctorDatesDialog> createState() => _DoctorDatesDialogState();
}

class _DoctorDatesDialogState extends ConsumerState<_DoctorDatesDialog> {
  String? _selectedShiftId;

  @override
  void initState() {
    super.initState();
    if (widget.shifts.isNotEmpty) {
      _selectedShiftId = widget.shifts.first.id;
    }
    WidgetsBinding.instance.addPostFrameCallback((_) {
      ref
          .read(rosterManagementControllerProvider.notifier)
          .loadDoctorDates(widget.doctor.id);
    });
  }

  Future<List<String>?> _pickIsoDates({required String title}) {
    final now = DateTime.now();
    var month = DateTime(now.year, now.month);
    final selected = <String>{};

    return showDialog<List<String>>(
      context: context,
      builder: (dialogContext) => StatefulBuilder(
        builder: (context, setDialogState) {
          final daysInMonth = DateUtils.getDaysInMonth(month.year, month.month);
          final monthLabel = DateFormat('MMMM yyyy').format(month);

          return AlertDialog(
            title: Text(title),
            content: SizedBox(
              width: 340,
              child: SingleChildScrollView(
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        IconButton(
                          tooltip: 'Previous month',
                          icon: const Icon(Icons.chevron_left),
                          onPressed: () => setDialogState(() {
                            month = DateTime(month.year, month.month - 1);
                          }),
                        ),
                        Text(
                          monthLabel,
                          style: Theme.of(context).textTheme.titleSmall
                              ?.copyWith(fontWeight: FontWeight.w700),
                        ),
                        IconButton(
                          tooltip: 'Next month',
                          icon: const Icon(Icons.chevron_right),
                          onPressed: () => setDialogState(() {
                            month = DateTime(month.year, month.month + 1);
                          }),
                        ),
                      ],
                    ),
                    const SizedBox(height: 8),
                    GridView.builder(
                      shrinkWrap: true,
                      physics: const NeverScrollableScrollPhysics(),
                      itemCount: daysInMonth,
                      gridDelegate:
                          const SliverGridDelegateWithFixedCrossAxisCount(
                            crossAxisCount: 7,
                            mainAxisSpacing: 6,
                            crossAxisSpacing: 6,
                            mainAxisExtent: 38,
                          ),
                    itemBuilder: (context, index) {
                      final day = index + 1;
                      final iso = DateFormat(
                        'yyyy-MM-dd',
                      ).format(DateTime(month.year, month.month, day));
                      final isSelected = selected.contains(iso);

                      return _SelectableTile(
                        label: '$day',
                        selected: isSelected,
                        onTap: () {
                          setDialogState(() {
                            if (isSelected) {
                              selected.remove(iso);
                            } else {
                              selected.add(iso);
                            }
                          });
                        },
                      );
                    },
                  ),
                  ],
                ),
              ),
            ),
            actions: [
              TextButton(
                onPressed: () => Navigator.of(dialogContext).pop(),
                child: const Text('Cancel'),
              ),
              FilledButton(
                onPressed: selected.isEmpty
                    ? null
                    : () {
                        final sorted = selected.toList()..sort();
                        Navigator.of(dialogContext).pop(sorted);
                      },
                child: Text(
                  selected.isEmpty
                      ? 'Add Dates'
                      : 'Add (${selected.length})',
                ),
              ),
            ],
          );
        },
      ),
    );
  }

  String _shiftLabel(String shiftId) {
    final shift = widget.shifts.where((s) => s.id == shiftId).firstOrNull;
    if (shift == null) return shiftId;
    final pos = widget.positions
        .where((p) => p.id == shift.positionId)
        .firstOrNull;
    return pos == null ? shift.name : '${pos.name} · ${shift.name}';
  }

  @override
  Widget build(BuildContext context) {
    final state = ref.watch(rosterManagementControllerProvider);
    final notifier = ref.read(rosterManagementControllerProvider.notifier);
    final scheme = Theme.of(context).colorScheme;

    return AlertDialog(
      title: Text('Schedule Dates — ${widget.doctor.name}'),
      content: SizedBox(
        width: 420,
        child: state.isLoadingDates
            ? const SizedBox(
                height: 160,
                child: Center(child: CircularProgressIndicator()),
              )
            : SingleChildScrollView(
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    if (state.errorMessage != null) ...[
                      Container(
                        width: double.infinity,
                        padding: const EdgeInsets.all(10),
                        margin: const EdgeInsets.only(bottom: 12),
                        decoration: BoxDecoration(
                          color: scheme.errorContainer,
                          borderRadius: BorderRadius.circular(8),
                        ),
                        child: Text(
                          state.errorMessage!,
                          style: TextStyle(color: scheme.onErrorContainer),
                        ),
                      ),
                    ],
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        Text(
                          'Unavailable Dates',
                          style: Theme.of(context).textTheme.titleSmall,
                        ),
                        TextButton.icon(
                          onPressed: state.isSaving
                              ? null
                              : () async {
                                  final dates = await _pickIsoDates(
                                    title: 'Select Unavailable Dates',
                                  );
                                  if (dates != null && dates.isNotEmpty) {
                                    await notifier.addDoctorUnavailabilities(
                                      doctorId: widget.doctor.id,
                                      dates: dates,
                                    );
                                  }
                                },
                          icon: const Icon(Icons.add, size: 18),
                          label: const Text('Add Dates'),
                        ),
                      ],
                    ),
                    if (state.selectedDoctorUnavailabilities.isEmpty)
                      const Text('No unavailable dates.')
                    else
                      Wrap(
                        spacing: 6,
                        runSpacing: 6,
                        children: [
                          for (final u in state.selectedDoctorUnavailabilities)
                            InputChip(
                              label: Text(u.date),
                              deleteButtonTooltipMessage:
                                  'Remove unavailability ${u.date}',
                              onDeleted: state.isSaving
                                  ? null
                                  : () => notifier.removeDoctorUnavailability(
                                      doctorId: widget.doctor.id,
                                      unavailabilityId: u.id,
                                    ),
                            ),
                        ],
                      ),
                    const Divider(height: 24),
                    Text(
                      'Pre-Assignments',
                      style: Theme.of(context).textTheme.titleSmall,
                    ),
                    const SizedBox(height: 8),
                    if (widget.shifts.isNotEmpty)
                      Row(
                        children: [
                          Expanded(
                            child: DropdownButtonFormField<String>(
                              initialValue: _selectedShiftId,
                              isExpanded: true,
                              decoration: const InputDecoration(
                                labelText: 'Shift',
                                isDense: true,
                                border: OutlineInputBorder(),
                              ),
                              items: [
                                for (final shift in widget.shifts)
                                  DropdownMenuItem<String>(
                                    value: shift.id,
                                    child: Text(
                                      _shiftLabel(shift.id),
                                      overflow: TextOverflow.ellipsis,
                                    ),
                                  ),
                              ],
                              onChanged: (v) =>
                                  setState(() => _selectedShiftId = v),
                            ),
                          ),
                          const SizedBox(width: 8),
                          FilledButton.tonal(
                            onPressed:
                                (state.isSaving || _selectedShiftId == null)
                                ? null
                                : () async {
                                    final dates = await _pickIsoDates(
                                      title: 'Select Pre-Assignment Dates',
                                    );
                                    if (dates != null && dates.isNotEmpty) {
                                      await notifier.addDoctorPreAssignments(
                                        doctorId: widget.doctor.id,
                                        shiftId: _selectedShiftId!,
                                        dates: dates,
                                      );
                                    }
                                  },
                            child: const Text('Assign Dates'),
                          ),
                        ],
                      ),
                    const SizedBox(height: 8),
                    if (state.selectedDoctorPreAssignments.isEmpty)
                      const Text('No pre-assignments.')
                    else
                      for (final pa in state.selectedDoctorPreAssignments)
                        ListTile(
                          dense: true,
                          contentPadding: EdgeInsets.zero,
                          title: Text('${pa.date} · ${_shiftLabel(pa.shiftId)}'),
                          trailing: IconButton(
                            tooltip: 'Remove pre-assignment ${pa.date}',
                            icon: const Icon(Icons.delete_outline, size: 20),
                            onPressed: state.isSaving
                                ? null
                                : () => notifier.removeDoctorPreAssignment(
                                    doctorId: widget.doctor.id,
                                    preAssignmentId: pa.id,
                                  ),
                          ),
                        ),
                  ],
                ),
              ),
      ),
      actions: [
        TextButton(
          onPressed: () => Navigator.of(context).pop(),
          child: const Text('Close'),
        ),
      ],
    );
  }
}
