import 'package:flutter/material.dart';

class AdminMetricCards extends StatelessWidget {
  const AdminMetricCards({
    super.key,
    this.dueDate = 'Oct 25',
    this.daysLeft = 5,
    this.submissionStatus = 'Requests Closed',
    this.pendingApprovalsCount = 2,
    this.pendingApprovalsStatus = 'Action Required',
    this.availableStaffCount = 18,
    this.totalShiftsCount = 90,
    this.staffCapacitySubtitle = '~5.0 shifts/doctor',
    this.bottleneckDaysCount = 3,
    this.bottleneckStatus = 'Requires Attention',
    this.bottleneckDates = const ['ICU - Nov 12', 'ER - Nov 15'],
  });

  final String dueDate;
  final int daysLeft;
  final String submissionStatus;
  final int pendingApprovalsCount;
  final String pendingApprovalsStatus;
  final int availableStaffCount;
  final int totalShiftsCount;
  final String staffCapacitySubtitle;
  final int bottleneckDaysCount;
  final String bottleneckStatus;
  final List<String> bottleneckDates;

  @override
  Widget build(BuildContext context) {
    return LayoutBuilder(
      builder: (context, constraints) {
        final card1 = _buildDueDateCard(context);
        final card2 = _buildPendingApprovalsCard(context);
        final card3 = _buildStaffPoolCard(context);
        final card4 = _buildBottlenecksCard(context);

        if (constraints.maxWidth > 1050) {
          // Desktop: 4 cards in a row
          return IntrinsicHeight(
            child: Row(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                Expanded(child: card1),
                const SizedBox(width: 12),
                Expanded(child: card2),
                const SizedBox(width: 12),
                Expanded(child: card3),
                const SizedBox(width: 12),
                Expanded(child: card4),
              ],
            ),
          );
        } else if (constraints.maxWidth > 650) {
          // Tablet: 2x2 grid
          return IntrinsicHeight(
            child: Row(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                Expanded(
                  child: Column(
                    children: [
                      Expanded(child: card1),
                      const SizedBox(height: 12),
                      Expanded(child: card3),
                    ],
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Column(
                    children: [
                      Expanded(child: card2),
                      const SizedBox(height: 12),
                      Expanded(child: card4),
                    ],
                  ),
                ),
              ],
            ),
          );
        } else {
          // Mobile: Stacked column
          return Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              card1,
              const SizedBox(height: 12),
              card2,
              const SizedBox(height: 12),
              card3,
              const SizedBox(height: 12),
              card4,
            ],
          );
        }
      },
    );
  }

  // Helper for the card shell
  Widget _buildCardShell(
    BuildContext context, {
    required List<Widget> children,
  }) {
    final theme = Theme.of(context);

    return SizedBox(
      width: double.infinity,
      child: Card(
        elevation: 0,
        shape: RoundedRectangleBorder(
          borderRadius: BorderRadius.circular(12),
          side: BorderSide(
            color: theme.colorScheme.outlineVariant.withValues(alpha: 0.6),
          ),
        ),
        child: Padding(
          padding: EdgeInsets.all(16),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: children,
          ),
        ),
      ),
    );
  }

  Widget _buildDueDateCard(BuildContext context) {
    final theme = Theme.of(context);

    return _buildCardShell(
      context,
      children: [
        Text(
          'Schedule Due Date',
          style: theme.textTheme.headlineMedium?.copyWith(
            fontWeight: FontWeight.bold,
          ),
        ),
        const SizedBox(height: 8),
        Text(
          dueDate,
          style: theme.textTheme.titleSmall?.copyWith(
            color: theme.colorScheme.onSurfaceVariant,
            fontWeight: FontWeight.w600,
            fontSize: 22,
          ),
        ),
        const SizedBox(height: 8),
        Row(
          children: [
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
              decoration: BoxDecoration(
                color: Colors.blue.shade50,
                borderRadius: BorderRadius.circular(12),
                border: Border.all(color: Colors.blue.shade200),
              ),
              child: Text(
                '$daysLeft Days Left',
                style: TextStyle(
                  color: Colors.blue.shade800,
                  fontSize: 11,
                  fontWeight: FontWeight.w600,
                ),
              ),
            ),
            const SizedBox(width: 8),
            Expanded(
              child: Text(
                submissionStatus,
                style: theme.textTheme.bodySmall?.copyWith(
                  color: theme.colorScheme.onSurfaceVariant,
                ),
                overflow: TextOverflow.ellipsis,
              ),
            ),
          ],
        ),
      ],
    );
  }

  Widget _buildPendingApprovalsCard(BuildContext context) {
    final theme = Theme.of(context);

    return _buildCardShell(
      context,
      children: [
        Text(
          'Pending Approvals',
          style: theme.textTheme.headlineMedium?.copyWith(
            fontWeight: FontWeight.bold,
          ),
        ),
        const SizedBox(height: 8),
        Row(
          mainAxisAlignment: MainAxisAlignment.spaceBetween,
          children: [
            Text(
              '$pendingApprovalsCount',
              style: theme.textTheme.titleSmall?.copyWith(
                color: theme.colorScheme.onSurfaceVariant,
                fontWeight: FontWeight.w600,
                fontSize: 22,
              ),
            ),
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
              decoration: BoxDecoration(
                color: Colors.red.shade50,
                borderRadius: BorderRadius.circular(12),
                border: Border.all(color: Colors.red.shade200),
              ),
              child: Text(
                pendingApprovalsStatus,
                style: TextStyle(
                  color: Colors.red.shade800,
                  fontSize: 11,
                  fontWeight: FontWeight.w600,
                ),
              ),
            ),
          ],
        ),
        const SizedBox(height: 8),
        Text(
          'Action required before solve',
          style: theme.textTheme.bodySmall?.copyWith(
            color: theme.colorScheme.onSurfaceVariant,
          ),
        ),
      ],
    );
  }

  Widget _buildStaffPoolCard(BuildContext context) {
    final theme = Theme.of(context);

    return _buildCardShell(
      context,
      children: [
        Text(
          'Staff Pool & Capacity',
          style: theme.textTheme.headlineMedium?.copyWith(
            fontWeight: FontWeight.bold,
          ),
        ),
        const SizedBox(height: 8),
        Text(
          '$availableStaffCount Doctors',
          style: theme.textTheme.titleSmall?.copyWith(
            color: theme.colorScheme.onSurfaceVariant,
            fontWeight: FontWeight.w600,
            fontSize: 22,
          ),
        ),
        const SizedBox(height: 8),
        Text(
          '$totalShiftsCount Shifts to Assign',
          style: theme.textTheme.bodyMedium?.copyWith(
            fontWeight: FontWeight.w500,
          ),
        ),
        const SizedBox(height: 2),
        Text(
          staffCapacitySubtitle,
          style: theme.textTheme.bodySmall?.copyWith(
            color: theme.colorScheme.onSurfaceVariant,
          ),
        ),
      ],
    );
  }

  Widget _buildBottlenecksCard(BuildContext context) {
    final theme = Theme.of(context);

    return _buildCardShell(
      context,
      children: [
        Text(
          'Coverage Bottlenecks',
          style: theme.textTheme.headlineMedium?.copyWith(
            fontWeight: FontWeight.bold,
          ),
        ),
        const SizedBox(height: 8),
        Row(
          mainAxisAlignment: MainAxisAlignment.spaceBetween,
          children: [
            Flexible(
              child: Text(
                '$bottleneckDaysCount Tight Days',
                style: theme.textTheme.titleSmall?.copyWith(
                  color: theme.colorScheme.onSurfaceVariant,
                  fontWeight: FontWeight.w600,
                  fontSize: 22,
                ),
              ),
            ),
            const SizedBox(width: 8),
            Container(
              padding: EdgeInsets.symmetric(horizontal: 8, vertical: 3),
              decoration: BoxDecoration(
                color: Colors.amber.shade50,
                borderRadius: BorderRadius.circular(12),
                border: Border.all(color: Colors.amber.shade300),
              ),
              child: Text(
                bottleneckStatus,
                style: TextStyle(
                  color: Colors.amber.shade900,
                  fontSize: 11,
                  fontWeight: FontWeight.w600,
                ),
              ),
            ),
          ],
        ),
        const SizedBox(height: 8),
        Wrap(
          spacing: 6,
          runSpacing: 4,
          children: bottleneckDates
              .map(
                (slot) => Container(
                  padding: const EdgeInsets.symmetric(
                    horizontal: 8,
                    vertical: 3,
                  ),
                  decoration: BoxDecoration(
                    color: theme.colorScheme.surfaceContainerHighest,
                    borderRadius: BorderRadius.circular(6),
                    border: Border.all(color: theme.colorScheme.outlineVariant),
                  ),
                  child: Text(
                    slot,
                    style: theme.textTheme.labelSmall?.copyWith(
                      fontWeight: FontWeight.w500,
                    ),
                  ),
                ),
              )
              .toList(),
        ),
      ],
    );
  }
}
