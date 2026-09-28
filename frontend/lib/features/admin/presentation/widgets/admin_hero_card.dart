import 'package:flutter/material.dart';

class AdminHeroCard extends StatelessWidget {
  const AdminHeroCard({
    super.key,
    required this.targetMonth,
    this.statusLabel = 'Ready to Generate',
    this.isGenerating = false,
    required this.onGeneratePressed,
  });

  final String targetMonth;
  final String statusLabel;
  final bool isGenerating;
  final VoidCallback? onGeneratePressed;

  @override
  Widget build(BuildContext context) {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(20),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                Expanded(
                  child: Text(
                    targetMonth,
                    style: TextStyle(fontWeight: FontWeight.bold, fontSize: 36),
                  ),
                ),
                Container(
                  padding: const EdgeInsets.symmetric(
                    horizontal: 10,
                    vertical: 4,
                  ),
                  decoration: BoxDecoration(
                    color: Colors.green.shade50,
                    borderRadius: BorderRadius.circular(16),
                    border: Border.all(color: Colors.green.shade200),
                  ),
                  child: Row(
                    mainAxisSize: MainAxisSize.min,
                    children: [
                      const Icon(Icons.circle, size: 8, color: Colors.green),
                      const SizedBox(width: 6),
                      Text(
                        statusLabel,
                        style: TextStyle(
                          color: Colors.green.shade800,
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
    );
  }
}
