import 'package:flutter/material.dart';

class DraftPreviewCanvas extends StatelessWidget {
  const DraftPreviewCanvas({
    super.key,
    this.isGenerated = false,
    this.generatedContent,
  });

  final bool isGenerated;
  final Widget? generatedContent;

  @override
  Widget build(BuildContext context) {
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
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Padding(
              padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
              child: Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  Text(
                    'Draft Preview Canvas',
                    style: theme.textTheme.titleMedium?.copyWith(
                      fontWeight: FontWeight.bold,
                    ),
                  ),
                  Row(
                    children: [
                      IconButton(
                        onPressed: () {},
                        icon: Icon(Icons.calendar_view_month),
                        tooltip: 'Calendar View',
                      ),
                      IconButton(
                        onPressed: () {},
                        icon: Icon(Icons.table_chart_outlined),
                        tooltip: 'Table View',
                      ),
                    ],
                  ),
                ],
              ),
            ),
            const Divider(height: 1),

            isGenerated
                ? Padding(
                    padding: const EdgeInsets.all(16),
                    child:
                        generatedContent ??
                        const Center(child: Text('Draft Schedule Generated')),
                  )
                : _buildEmptyState(context),
          ],
        ),
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
}
