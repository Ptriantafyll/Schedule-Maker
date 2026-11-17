import 'package:flutter/gestures.dart';
import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/admin/domain/models/roster_doctor.dart';
import 'package:frontend/features/admin/domain/models/roster_position.dart';
import 'package:frontend/features/admin/domain/models/roster_shift.dart';
import 'package:frontend/features/admin/domain/models/roster_team.dart';
import 'package:frontend/features/admin/presentation/controllers/roster_management_controller.dart';
import 'package:frontend/features/admin/presentation/widgets/roster_dialogs.dart';

class RosterManagementScreen extends ConsumerStatefulWidget {
  const RosterManagementScreen({super.key});

  @override
  ConsumerState<RosterManagementScreen> createState() =>
      _RosterManagementScreenState();
}

class _RosterManagementScreenState
    extends ConsumerState<RosterManagementScreen> {
  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addPostFrameCallback((_) {
      ref.read(rosterManagementControllerProvider.notifier).loadRoster();
    });
  }

  void _showSuccess(String message) {
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text(message)),
    );
  }

  Future<void> _confirmDelete({
    required String title,
    required String message,
    required String successText,
    required Future<bool> Function(RosterManagementController notifier) onDelete,
  }) async {
    final confirmed = await showRosterConfirmDeleteDialog(
      context,
      title: title,
      message: message,
    );
    if (!confirmed || !mounted) return;

    final ok = await onDelete(
      ref.read(rosterManagementControllerProvider.notifier),
    );
    if (ok) _showSuccess(successText);
  }

  Future<void> _handleAddOrEditTeam({RosterTeam? existing}) async {
    final name = await showRosterTeamDialog(context, initialTeam: existing);
    if (name == null || !mounted) return;

    final notifier = ref.read(rosterManagementControllerProvider.notifier);
    final ok = existing == null
        ? await notifier.createTeam(name: name)
        : await notifier.updateTeam(teamId: existing.id, name: name);
    if (ok) {
      _showSuccess(existing == null ? 'Team created.' : 'Team updated.');
    }
  }

  Future<void> _handleAddOrEditPosition({RosterPosition? existing}) async {
    final result = await showRosterPositionDialog(
      context,
      initialPosition: existing,
    );
    if (result == null || !mounted) return;

    final notifier = ref.read(rosterManagementControllerProvider.notifier);
    final ok = existing == null
        ? await notifier.createPosition(
            name: result.name,
            dutyDays: result.dutyDays,
          )
        : await notifier.updatePosition(
            positionId: existing.id,
            name: result.name,
            dutyDays: result.dutyDays,
          );
    if (ok) {
      _showSuccess(
        existing == null ? 'Position created.' : 'Position updated.',
      );
    }
  }

  Future<void> _handleAddOrEditShift({
    required String positionId,
    RosterShift? existing,
  }) async {
    final result = await showRosterShiftDialog(context, initialShift: existing);
    if (result == null || !mounted) return;

    final notifier = ref.read(rosterManagementControllerProvider.notifier);
    final ok = existing == null
        ? await notifier.createShift(
            positionId: positionId,
            name: result.name,
            doctorsPerShift: result.doctorsPerShift,
            grantsDayOff: result.grantsDayOff,
          )
        : await notifier.updateShift(
            shiftId: existing.id,
            name: result.name,
            doctorsPerShift: result.doctorsPerShift,
            grantsDayOff: result.grantsDayOff,
          );
    if (ok) {
      _showSuccess(existing == null ? 'Shift created.' : 'Shift updated.');
    }
  }

  Future<void> _handleAddOrEditDoctor({
    RosterDoctor? existing,
    required List<RosterTeam> teams,
    required List<RosterPosition> positions,
  }) async {
    final result = await showRosterDoctorDialog(
      context,
      initialDoctor: existing,
      teams: teams,
      positions: positions,
    );
    if (result == null || !mounted) return;

    final notifier = ref.read(rosterManagementControllerProvider.notifier);
    final ok = existing == null
        ? await notifier.createDoctor(
            name: result.name,
            teamId: result.teamId,
            positionIds: result.positionIds,
          )
        : await notifier.updateDoctor(
            existingDoctor: existing,
            name: result.name,
            teamId: result.teamId,
            positionIds: result.positionIds,
          );
    if (ok) {
      _showSuccess(existing == null ? 'Doctor added.' : 'Doctor updated.');
    }
  }

  @override
  Widget build(BuildContext context) {
    ref.listen<RosterManagementState>(rosterManagementControllerProvider, (
      prev,
      next,
    ) {
      if (next.errorMessage != null &&
          next.errorMessage != prev?.errorMessage) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text(next.errorMessage!),
            backgroundColor: Theme.of(context).colorScheme.error,
          ),
        );
      }
    });

    final state = ref.watch(rosterManagementControllerProvider);
    final teamNames = {for (final t in state.teams) t.id: t.name};
    final posNames = {for (final p in state.positions) p.id: p.name};

    return ScrollConfiguration(
      behavior: ScrollConfiguration.of(context).copyWith(
        dragDevices: {
          PointerDeviceKind.touch,
          PointerDeviceKind.mouse,
          PointerDeviceKind.trackpad,
        },
      ),
      child: Scaffold(
        appBar: AppBar(
          title: const Text('Department Roster'),
          actions: [
            IconButton(
              tooltip: 'Refresh roster',
              onPressed: state.isLoading
                  ? null
                  : () => ref
                        .read(rosterManagementControllerProvider.notifier)
                        .loadRoster(),
              icon: const Icon(Icons.refresh),
            ),
          ],
        ),
        body: state.isLoading
            ? const Center(child: CircularProgressIndicator())
            : RefreshIndicator(
                onRefresh: () => ref
                    .read(rosterManagementControllerProvider.notifier)
                    .loadRoster(),
                child: ListView(
                  padding: const EdgeInsets.symmetric(vertical: 16),
                  children: [
                    _OverviewBanner(state: state),
                    const SizedBox(height: 20),
                    _SectionHeader(
                      title: 'Doctors',
                      count: state.doctors.length,
                      actionLabel: 'Add Doctor',
                      onAction: () => _handleAddOrEditDoctor(
                        teams: state.teams,
                        positions: state.positions,
                      ),
                    ),
                    const SizedBox(height: 10),
                    _HorizontalStrip(
                      height: 224,
                      itemCount: state.doctors.length,
                      emptyMessage: 'No doctors added to this department yet.',
                      itemBuilder: (context, index) {
                        final doctor = state.doctors[index];
                        return _DoctorCard(
                          doctor: doctor,
                          teamName: doctor.teamId != null
                              ? teamNames[doctor.teamId]
                              : null,
                          positionNames: [
                            for (final pid in doctor.positionIds)
                              if (posNames[pid] != null) posNames[pid]!,
                          ],
                          onDatesPressed: () => showDoctorDatesDialog(
                            context,
                            doctor: doctor,
                            shifts: state.shifts,
                            positions: state.positions,
                          ),
                          onEditPressed: () => _handleAddOrEditDoctor(
                            existing: doctor,
                            teams: state.teams,
                            positions: state.positions,
                          ),
                          onDeletePressed: () => _confirmDelete(
                            title: 'Remove Doctor',
                            message:
                                'Remove "${doctor.name}" from the department roster? Any linked user account remains active.',
                            successText: 'Doctor removed.',
                            onDelete: (n) => n.deleteDoctor(doctorId: doctor.id),
                          ),
                        );
                      },
                    ),
                    const SizedBox(height: 24),
                    _SectionHeader(
                      title: 'Positions & Shifts',
                      count: state.positions.length,
                      actionLabel: 'Add Position',
                      onAction: () => _handleAddOrEditPosition(),
                    ),
                    const SizedBox(height: 10),
                    _HorizontalStrip(
                      height: 290,
                      itemCount: state.positions.length,
                      emptyMessage:
                          'Add a position to define shifts and duty days.',
                      itemBuilder: (context, index) {
                        final position = state.positions[index];
                        final posShifts = state.shifts
                            .where((s) => s.positionId == position.id)
                            .toList();
                        return _PositionCard(
                          position: position,
                          shifts: posShifts,
                          onEditPosition: () =>
                              _handleAddOrEditPosition(existing: position),
                          onDeletePosition: () => _confirmDelete(
                            title: 'Delete Position',
                            message:
                                'Delete "${position.name}"? This also removes its shifts, doctor eligibility links, and upcoming pre-assignments.',
                            successText: 'Position deleted.',
                            onDelete: (n) =>
                                n.deletePosition(positionId: position.id),
                          ),
                          onAddShift: () =>
                              _handleAddOrEditShift(positionId: position.id),
                          onEditShift: (shift) => _handleAddOrEditShift(
                            positionId: position.id,
                            existing: shift,
                          ),
                          onDeleteShift: (shift) => _confirmDelete(
                            title: 'Delete Shift',
                            message:
                                'Delete "${shift.name}"? Upcoming pre-assignments for this shift will also be removed.',
                            successText: 'Shift deleted.',
                            onDelete: (n) => n.deleteShift(shiftId: shift.id),
                          ),
                        );
                      },
                    ),
                    const SizedBox(height: 24),
                    _SectionHeader(
                      title: 'Teams',
                      count: state.teams.length,
                      actionLabel: 'Add Team',
                      onAction: () => _handleAddOrEditTeam(),
                    ),
                    const SizedBox(height: 10),
                    _HorizontalStrip(
                      height: 156,
                      itemCount: state.teams.length,
                      emptyMessage:
                          'No teams yet. Doctors without a team are scheduled individually.',
                      itemBuilder: (context, index) {
                        final team = state.teams[index];
                        final count = state.doctors
                            .where((d) => d.teamId == team.id)
                            .length;
                        return _TeamCard(
                          team: team,
                          doctorCount: count,
                          onRenamePressed: () =>
                              _handleAddOrEditTeam(existing: team),
                          onDeletePressed: () => _confirmDelete(
                            title: 'Delete Team',
                            message:
                                'Delete "${team.name}"? Assigned doctors will remain active without a team.',
                            successText: 'Team deleted.',
                            onDelete: (n) => n.deleteTeam(teamId: team.id),
                          ),
                        );
                      },
                    ),
                    const SizedBox(height: 24),
                  ],
                ),
              ),
      ),
    );
  }
}

class _OverviewBanner extends StatelessWidget {
  const _OverviewBanner({required this.state});
  final RosterManagementState state;

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 20),
      child: Container(
        padding: const EdgeInsets.all(16),
        decoration: BoxDecoration(
          color: scheme.surfaceContainerLow,
          borderRadius: BorderRadius.circular(16),
          border: Border.all(color: scheme.outlineVariant),
        ),
        child: Wrap(
          spacing: 12,
          runSpacing: 8,
          crossAxisAlignment: WrapCrossAlignment.center,
          children: [
            _MetricPill(
              icon: Icons.medical_services_outlined,
              label: '${state.doctors.length} Doctors',
            ),
            _MetricPill(
              icon: Icons.work_outline,
              label: '${state.positions.length} Positions',
            ),
            _MetricPill(
              icon: Icons.schedule_outlined,
              label: '${state.shifts.length} Shifts',
            ),
            _MetricPill(
              icon: Icons.groups_outlined,
              label: '${state.teams.length} Teams',
            ),
          ],
        ),
      ),
    );
  }
}

class _MetricPill extends StatelessWidget {
  const _MetricPill({required this.icon, required this.label});
  final IconData icon;
  final String label;

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
      decoration: BoxDecoration(
        color: scheme.surface,
        borderRadius: BorderRadius.circular(20),
        border: Border.all(color: scheme.outlineVariant),
      ),
      child: Row(
        mainAxisSize: MainAxisSize.min,
        children: [
          Icon(icon, size: 16, color: scheme.primary),
          const SizedBox(width: 6),
          Text(label, style: Theme.of(context).textTheme.labelMedium),
        ],
      ),
    );
  }
}

class _SectionHeader extends StatelessWidget {
  const _SectionHeader({
    required this.title,
    required this.count,
    required this.actionLabel,
    required this.onAction,
  });

  final String title;
  final int count;
  final String actionLabel;
  final VoidCallback onAction;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 20),
      child: Row(
        children: [
          Text(
            title,
            style: theme.textTheme.titleMedium?.copyWith(
              fontWeight: FontWeight.w700,
            ),
          ),
          const SizedBox(width: 8),
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 2),
            decoration: BoxDecoration(
              color: theme.colorScheme.secondaryContainer,
              borderRadius: BorderRadius.circular(12),
            ),
            child: Text(
              '$count',
              style: theme.textTheme.labelSmall?.copyWith(
                color: theme.colorScheme.onSecondaryContainer,
                fontWeight: FontWeight.w700,
              ),
            ),
          ),
          const Spacer(),
          FilledButton.icon(
            onPressed: onAction,
            icon: const Icon(Icons.add, size: 18),
            label: Text(actionLabel),
          ),
        ],
      ),
    );
  }
}

class _HorizontalStrip extends StatelessWidget {
  const _HorizontalStrip({
    required this.height,
    required this.itemCount,
    required this.emptyMessage,
    required this.itemBuilder,
  });

  final double height;
  final int itemCount;
  final String emptyMessage;
  final IndexedWidgetBuilder itemBuilder;

  @override
  Widget build(BuildContext context) {
    if (itemCount == 0) {
      return Padding(
        padding: const EdgeInsets.symmetric(horizontal: 20),
        child: Card(
          elevation: 0,
          shape: RoundedRectangleBorder(
            borderRadius: BorderRadius.circular(16),
            side: BorderSide(
              color: Theme.of(context).colorScheme.outlineVariant,
            ),
          ),
          child: Padding(
            padding: const EdgeInsets.all(20),
            child: Text(
              emptyMessage,
              style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                color: Theme.of(context).colorScheme.onSurfaceVariant,
              ),
            ),
          ),
        ),
      );
    }

    return SizedBox(
      height: height,
      child: ListView.separated(
        scrollDirection: Axis.horizontal,
        padding: const EdgeInsets.symmetric(horizontal: 20),
        itemCount: itemCount,
        separatorBuilder: (_, _) => const SizedBox(width: 12),
        itemBuilder: itemBuilder,
      ),
    );
  }
}

class _DoctorCard extends StatelessWidget {
  const _DoctorCard({
    required this.doctor,
    required this.teamName,
    required this.positionNames,
    required this.onDatesPressed,
    required this.onEditPressed,
    required this.onDeletePressed,
  });

  final RosterDoctor doctor;
  final String? teamName;
  final List<String> positionNames;
  final VoidCallback onDatesPressed;
  final VoidCallback onEditPressed;
  final VoidCallback onDeletePressed;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final scheme = theme.colorScheme;
    final initial = doctor.name.isNotEmpty
        ? doctor.name.characters.first.toUpperCase()
        : '?';

    return SizedBox(
      width: 296,
      child: Card(
        elevation: 0,
        shape: RoundedRectangleBorder(
          borderRadius: BorderRadius.circular(16),
          side: BorderSide(color: scheme.outlineVariant),
        ),
        child: Padding(
          padding: const EdgeInsets.all(14),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Row(
                children: [
                  CircleAvatar(
                    radius: 18,
                    backgroundColor: scheme.primaryContainer,
                    child: Text(
                      initial,
                      style: TextStyle(
                        color: scheme.onPrimaryContainer,
                        fontWeight: FontWeight.w700,
                      ),
                    ),
                  ),
                  const SizedBox(width: 10),
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(
                          doctor.name,
                          maxLines: 1,
                          overflow: TextOverflow.ellipsis,
                          style: theme.textTheme.titleSmall?.copyWith(
                            fontWeight: FontWeight.w700,
                          ),
                        ),
                        const SizedBox(height: 2),
                        Container(
                          padding: const EdgeInsets.symmetric(
                            horizontal: 8,
                            vertical: 2,
                          ),
                          decoration: BoxDecoration(
                            color: teamName != null
                                ? scheme.secondaryContainer
                                : scheme.surfaceContainerHighest,
                            borderRadius: BorderRadius.circular(8),
                          ),
                          child: Text(
                            teamName ?? 'No Team',
                            style: theme.textTheme.labelSmall?.copyWith(
                              color: teamName != null
                                  ? scheme.onSecondaryContainer
                                  : scheme.onSurfaceVariant,
                            ),
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 12),
              Expanded(
                child: positionNames.isEmpty
                    ? Text(
                        'No positions assigned',
                        style: theme.textTheme.bodySmall?.copyWith(
                          color: scheme.onSurfaceVariant,
                        ),
                      )
                    : SingleChildScrollView(
                        child: Wrap(
                          spacing: 6,
                          runSpacing: 6,
                          children: [
                            for (final posName in positionNames)
                              Container(
                                padding: const EdgeInsets.symmetric(
                                  horizontal: 8,
                                  vertical: 4,
                                ),
                                decoration: BoxDecoration(
                                  color: scheme.surfaceContainerHighest,
                                  borderRadius: BorderRadius.circular(8),
                                ),
                                child: Text(
                                  posName,
                                  style: theme.textTheme.labelSmall,
                                ),
                              ),
                          ],
                        ),
                      ),
              ),
              Row(
                children: [
                  OutlinedButton.icon(
                    onPressed: onDatesPressed,
                    icon: const Icon(Icons.event_note_outlined, size: 16),
                    label: const Text('Dates'),
                  ),
                  const Spacer(),
                  IconButton(
                    tooltip: 'Edit ${doctor.name}',
                    icon: const Icon(Icons.edit_outlined),
                    onPressed: onEditPressed,
                  ),
                  IconButton(
                    tooltip: 'Delete ${doctor.name}',
                    icon: Icon(Icons.delete_outline, color: scheme.error),
                    onPressed: onDeletePressed,
                  ),
                ],
              ),
            ],
          ),
        ),
      ),
    );
  }
}

class _PositionCard extends StatelessWidget {
  const _PositionCard({
    required this.position,
    required this.shifts,
    required this.onEditPosition,
    required this.onDeletePosition,
    required this.onAddShift,
    required this.onEditShift,
    required this.onDeleteShift,
  });

  final RosterPosition position;
  final List<RosterShift> shifts;
  final VoidCallback onEditPosition;
  final VoidCallback onDeletePosition;
  final VoidCallback onAddShift;
  final ValueChanged<RosterShift> onEditShift;
  final ValueChanged<RosterShift> onDeleteShift;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final scheme = theme.colorScheme;

    return SizedBox(
      width: 328,
      child: Card(
        elevation: 0,
        shape: RoundedRectangleBorder(
          borderRadius: BorderRadius.circular(16),
          side: BorderSide(color: scheme.outlineVariant),
        ),
        child: Padding(
          padding: const EdgeInsets.all(14),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Row(
                children: [
                  Expanded(
                    child: Text(
                      position.name,
                      maxLines: 1,
                      overflow: TextOverflow.ellipsis,
                      style: theme.textTheme.titleSmall?.copyWith(
                        fontWeight: FontWeight.w700,
                      ),
                    ),
                  ),
                  IconButton(
                    tooltip: 'Edit ${position.name}',
                    icon: const Icon(Icons.edit_outlined, size: 20),
                    onPressed: onEditPosition,
                  ),
                  IconButton(
                    tooltip: 'Delete ${position.name}',
                    icon: Icon(
                      Icons.delete_outline,
                      size: 20,
                      color: scheme.error,
                    ),
                    onPressed: onDeletePosition,
                  ),
                ],
              ),
              Wrap(
                spacing: 4,
                runSpacing: 4,
                children: [
                  for (final day in position.dutyDays)
                    if (day >= 0 && day < kWeekdayLabels.length)
                      Container(
                        padding: const EdgeInsets.symmetric(
                          horizontal: 7,
                          vertical: 2,
                        ),
                        decoration: BoxDecoration(
                          color: scheme.primaryContainer,
                          borderRadius: BorderRadius.circular(6),
                        ),
                        child: Text(
                          kWeekdayLabels[day],
                          style: theme.textTheme.labelSmall?.copyWith(
                            color: scheme.onPrimaryContainer,
                            fontWeight: FontWeight.w600,
                          ),
                        ),
                      ),
                ],
              ),
              const SizedBox(height: 10),
              Expanded(
                child: shifts.isEmpty
                    ? Center(
                        child: Text(
                          'No shifts configured',
                          style: theme.textTheme.bodySmall?.copyWith(
                            color: scheme.onSurfaceVariant,
                          ),
                        ),
                      )
                    : ListView.separated(
                        itemCount: shifts.length,
                        separatorBuilder: (_, _) => const SizedBox(height: 6),
                        itemBuilder: (context, index) {
                          final shift = shifts[index];
                          return Container(
                            padding: const EdgeInsets.symmetric(
                              horizontal: 10,
                              vertical: 6,
                            ),
                            decoration: BoxDecoration(
                              color: scheme.surfaceContainerLow,
                              borderRadius: BorderRadius.circular(10),
                            ),
                            child: Row(
                              children: [
                                Expanded(
                                  child: Column(
                                    crossAxisAlignment:
                                        CrossAxisAlignment.start,
                                    children: [
                                      Text(
                                        shift.name,
                                        style: theme.textTheme.bodyMedium
                                            ?.copyWith(
                                              fontWeight: FontWeight.w600,
                                            ),
                                      ),
                                      Text(
                                        '${shift.doctorsPerShift} doc(s)${shift.grantsDayOff ? ' · Day off' : ''}',
                                        style: theme.textTheme.labelSmall
                                            ?.copyWith(
                                              color: scheme.onSurfaceVariant,
                                            ),
                                      ),
                                    ],
                                  ),
                                ),
                                IconButton(
                                  tooltip: 'Edit ${shift.name}',
                                  icon: const Icon(
                                    Icons.edit_outlined,
                                    size: 18,
                                  ),
                                  onPressed: () => onEditShift(shift),
                                ),
                                IconButton(
                                  tooltip: 'Delete ${shift.name}',
                                  icon: Icon(
                                    Icons.delete_outline,
                                    size: 18,
                                    color: scheme.error,
                                  ),
                                  onPressed: () => onDeleteShift(shift),
                                ),
                              ],
                            ),
                          );
                        },
                      ),
              ),
              const SizedBox(height: 8),
              SizedBox(
                width: double.infinity,
                child: OutlinedButton.icon(
                  onPressed: onAddShift,
                  icon: const Icon(Icons.add, size: 18),
                  label: const Text('Add Shift'),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}

class _TeamCard extends StatelessWidget {
  const _TeamCard({
    required this.team,
    required this.doctorCount,
    required this.onRenamePressed,
    required this.onDeletePressed,
  });

  final RosterTeam team;
  final int doctorCount;
  final VoidCallback onRenamePressed;
  final VoidCallback onDeletePressed;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final scheme = theme.colorScheme;
    final countLabel =
        doctorCount == 1 ? '1 doctor' : '$doctorCount doctors';

    return SizedBox(
      width: 248,
      child: Card(
        elevation: 0,
        shape: RoundedRectangleBorder(
          borderRadius: BorderRadius.circular(16),
          side: BorderSide(color: scheme.outlineVariant),
        ),
        child: Padding(
          padding: const EdgeInsets.all(14),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Row(
                children: [
                  Container(
                    padding: const EdgeInsets.all(8),
                    decoration: BoxDecoration(
                      color: scheme.primaryContainer,
                      borderRadius: BorderRadius.circular(10),
                    ),
                    child: Icon(
                      Icons.groups_outlined,
                      size: 20,
                      color: scheme.onPrimaryContainer,
                    ),
                  ),
                  const SizedBox(width: 10),
                  Expanded(
                    child: Text(
                      team.name,
                      maxLines: 1,
                      overflow: TextOverflow.ellipsis,
                      style: theme.textTheme.titleSmall?.copyWith(
                        fontWeight: FontWeight.w700,
                      ),
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 10),
              Text(
                countLabel,
                style: theme.textTheme.bodySmall?.copyWith(
                  color: scheme.onSurfaceVariant,
                ),
              ),
              const Spacer(),
              Row(
                mainAxisAlignment: MainAxisAlignment.end,
                children: [
                  IconButton(
                    tooltip: 'Rename ${team.name}',
                    icon: const Icon(Icons.edit_outlined),
                    onPressed: onRenamePressed,
                  ),
                  IconButton(
                    tooltip: 'Delete ${team.name}',
                    icon: Icon(Icons.delete_outline, color: scheme.error),
                    onPressed: onDeletePressed,
                  ),
                ],
              ),
            ],
          ),
        ),
      ),
    );
  }
}
