import 'package:flutter/material.dart';

enum GenerationMode { currentRoster, excelUpload }

class GenerationModeDialog extends StatelessWidget {
  const GenerationModeDialog({super.key});

  @override
  Widget build(BuildContext context) {
    return AlertDialog(
      title: const Text('Select Generation Mode'),
      content: Column(
        mainAxisSize: MainAxisSize.min,
        children: [
          ListTile(
            leading: const Icon(Icons.people_alt_outlined),
            title: const Text('Current Department Roster'),
            subtitle: const Text(
              'Generate using in-app doctors and approved unavailabilities.',
            ),
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(8),
            ),
            onTap: () {
              Navigator.of(context).pop(GenerationMode.currentRoster);
            },
          ),
          const SizedBox(height: 8),

          ListTile(
            leading: const Icon(Icons.upload_file_outlined),
            title: const Text('Import from Excel (.xlsx)'),
            subtitle: const Text(
              'Upload a workbook containing doctor rosters and unavailabilities.',
            ),
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(8),
            ),
            onTap: () {
              Navigator.of(context).pop(GenerationMode.excelUpload);
            },
          ),
        ],
      ),
      actions: [
        TextButton(
          onPressed: () {
            Navigator.of(context).pop(); 
          },
          child: const Text('Cancel'),
        ),
      ],
    );
  }
}
