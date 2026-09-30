import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';

class ExcelUploadDialog extends StatefulWidget {
  const ExcelUploadDialog({super.key, this.onPickFile});

  final Future<PlatformFile?> Function()? onPickFile;

  @override
  State<StatefulWidget> createState() => _ExcelUploadDialogState();
}

class _ExcelUploadDialogState extends State<ExcelUploadDialog> {
  PlatformFile? _selectedFile;

  @override
  Widget build(BuildContext context) {
    return AlertDialog(
      title: const Row(
        children: [
          Icon(Icons.table_view_outlined),
          SizedBox(width: 8),
          Text('Import Schedule Workbook'),
        ],
      ),
      content: Column(
        mainAxisSize: MainAxisSize.min,
        crossAxisAlignment: CrossAxisAlignment.stretch,
        children: [
          Text(
            'Select an Excel (.xlsx) file containing doctor rosters and unavailabilities.',
            style: Theme.of(context).textTheme.bodySmall,
          ),
          const SizedBox(height: 16),

          if (_selectedFile == null)
            InkWell(
              onTap: _handlePickFile,
              borderRadius: BorderRadius.circular(8),
              child: Container(
                padding: const EdgeInsets.symmetric(
                  vertical: 24,
                  horizontal: 16,
                ),
                decoration: BoxDecoration(
                  border: Border.all(
                    color: Theme.of(context).colorScheme.outlineVariant,
                  ),
                  borderRadius: BorderRadius.circular(8),
                ),
                child: const Column(
                  children: [
                    Icon(Icons.cloud_upload_outlined, size: 36),
                    SizedBox(height: 8),
                    Text(
                      'Click to select .xlsx file',
                      style: TextStyle(fontWeight: FontWeight.w500),
                    ),
                  ],
                ),
              ),
            )
          else
            // File is selected, show file name, size, and close icon
            Container(
              padding: const EdgeInsets.all(12),
              decoration: BoxDecoration(
                border: Border.all(
                  color: Theme.of(context).colorScheme.primary,
                ),
                borderRadius: BorderRadius.circular(8),
              ),
              child: Row(
                children: [
                  const Icon(Icons.description_outlined),
                  const SizedBox(width: 12),
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(
                          _selectedFile!.name,
                          style: const TextStyle(fontWeight: FontWeight.bold),
                          maxLines: 1,
                          overflow: TextOverflow.ellipsis,
                        ),
                        Text(
                          _formatFileSize(_selectedFile!.lengthSync() ?? 0),
                          style: Theme.of(context).textTheme.bodySmall,
                        ),
                      ],
                    ),
                  ),
                  IconButton(
                    icon: const Icon(Icons.close),
                    onPressed: () => setState(() => _selectedFile = null),
                  ),
                ],
              ),
            ),
        ],
      ),
      actions: [
        TextButton(
          onPressed: () => Navigator.of(context).pop(),
          child: const Text('Cancel'),
        ),
        FilledButton(
          onPressed: _selectedFile == null
              ? null
              : () => Navigator.of(context).pop(_selectedFile),
          child: const Text('Generate Schedule'),
        ),
      ],
    );
  }

  Future<void> _handlePickFile() async {
    final file = widget.onPickFile != null
        ? await widget.onPickFile!()
        : await FilePicker.pickFile(
            type: FileType.custom,
            allowedExtensions: ['xlsx'],
          );

    if (file != null && mounted) {
      setState(() => _selectedFile = file);
    }
  }

  String _formatFileSize(int bytes) {
    return '${(bytes / 1024).toStringAsFixed(1)} KB';
  }
}
