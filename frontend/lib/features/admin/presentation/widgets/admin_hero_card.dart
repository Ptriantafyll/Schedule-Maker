import 'package:flutter/material.dart';
import 'package:skeletonizer/skeletonizer.dart';

class AdminHeroCard extends StatelessWidget {
  const AdminHeroCard({
    super.key,
    required this.targetMonth,
    this.statusLabel = 'Draft',
    this.isGenerating = false,
    required this.onGeneratePressed,
    this.isPublished = false,
    this.availableMonths = const [],
    this.onMonthSelected,
    this.isLoading = false,
  });

  final String targetMonth;
  final String statusLabel;
  final bool isGenerating;
  final VoidCallback? onGeneratePressed;
  final bool isPublished;
  final List<String> availableMonths;
  final ValueChanged<String>? onMonthSelected;
  final bool isLoading;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final badgeLabel = isPublished ? 'Published' : statusLabel;
    final badgeBgColor = isLoading
        ? theme.colorScheme.surfaceContainerHighest
        : (isPublished ? Colors.green.shade50 : Colors.amber.shade50);
    final badgeBorderColor = isLoading
        ? theme.colorScheme.outlineVariant
        : (isPublished ? Colors.green.shade200 : Colors.amber.shade200);
    final badgeDotColor = isLoading
        ? theme.colorScheme.outlineVariant
        : (isPublished ? Colors.green : Colors.amber.shade700);
    final badgeTextColor = isLoading
        ? theme.colorScheme.onSurfaceVariant
        : (isPublished ? Colors.green.shade800 : Colors.amber.shade900);

    final displayMonth = targetMonth.isNotEmpty ? targetMonth : 'November 2026';

    final hasMonthSelector =
        !isLoading && availableMonths.isNotEmpty && onMonthSelected != null;

    Widget monthHeader = Text(
      displayMonth,
      style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 36),
    );

    if (hasMonthSelector) {
      monthHeader = PopupMenuButton<String>(
        tooltip: 'Select target month',
        initialValue: targetMonth,
        onSelected: onMonthSelected,
        itemBuilder: (context) {
          return availableMonths.map((month) {
            return PopupMenuItem<String>(
              value: month,
              child: Row(
                children: [
                  Text(month),
                  if (month == targetMonth) ...[
                    const Spacer(),
                    const Icon(Icons.check, size: 18),
                  ],
                ],
              ),
            );
          }).toList();
        },
        child: Row(
          mainAxisSize: MainAxisSize.min,
          children: [
            Flexible(child: monthHeader),
            const SizedBox(width: 4),
            const Icon(Icons.arrow_drop_down, size: 36),
          ],
        ),
      );
    }

    return Skeletonizer(
      enabled: isLoading,
      child: Card(
        child: Padding(
          padding: const EdgeInsets.all(20),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  Flexible(child: monthHeader),
                  Container(
                    padding: const EdgeInsets.symmetric(
                      horizontal: 10,
                      vertical: 4,
                    ),
                    decoration: BoxDecoration(
                      color: badgeBgColor,
                      borderRadius: BorderRadius.circular(16),
                      border: Border.all(color: badgeBorderColor),
                    ),
                    child: Row(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        Icon(Icons.circle, size: 8, color: badgeDotColor),
                        const SizedBox(width: 6),
                        Text(
                          badgeLabel,
                          style: TextStyle(
                            color: badgeTextColor,
                            fontSize: 12,
                            fontWeight: FontWeight.w600,
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 12),
              Text(
                'Algorithm is primed with all staff requests, leave approvals, and coverage requirements. Estimated processing time: ~45 seconds.',
                style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                  color: Theme.of(context).colorScheme.onSurfaceVariant,
                ),
              ),

              const SizedBox(height: 20),
              FilledButton.icon(
                onPressed: isGenerating ? null : onGeneratePressed,
                icon: isGenerating
                    ? const SizedBox(
                        width: 18,
                        height: 18,
                        child: CircularProgressIndicator(
                          strokeWidth: 2,
                          color: Colors.white,
                        ),
                      )
                    : const Icon(Icons.auto_awesome),
                label: const Text('Generate Schedule'),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
