import 'dart:typed_data';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/admin/presentation/controllers/schedule_generation_controller.dart';
import 'package:frontend/features/admin/presentation/widgets/admin_hero_card.dart';
import 'package:frontend/features/admin/presentation/widgets/admin_metric_cards.dart';
import 'package:frontend/features/admin/presentation/widgets/draft_preview_canvas.dart';
import 'package:frontend/features/admin/presentation/widgets/excel_upload_dialog.dart';
import 'package:frontend/features/admin/presentation/widgets/generation_mode_dialog.dart';
import 'package:frontend/features/auth/presentation/controllers/auth_controller.dart';
import 'package:frontend/shared/widgets/bottom_nav_bar.dart';
import 'package:frontend/shared/widgets/profile_drawer.dart';

class AdminScreen extends ConsumerStatefulWidget {
  const AdminScreen({super.key, this.onSaveFile});

  final Future<String?> Function({
    required String fileName,
    required List<int> bytes,
  })? onSaveFile;

  @override
  ConsumerState<AdminScreen> createState() => _AdminScreenState();
}

class _AdminScreenState extends ConsumerState<AdminScreen> {
  int _currentIndex = 3;

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addPostFrameCallback((_) {
      _loadActiveDraft();
    });
  }

  void _loadActiveDraft() {
    final user = ref.read(currentUserProvider);
    ref.read(scheduleGenerationControllerProvider.notifier).loadActiveDraft(
      targetMonth: '2026-11',
      departmentId: user?.departmentId,
    );
  }

  Future<void> _handleGeneratePressed() async {
    final mode = await showDialog<GenerationMode>(
      context: context,
      builder: (_) => const GenerationModeDialog(),
    );

    if (!mounted || mode == null) return;

    if (mode == GenerationMode.currentRoster) {
      await _generateFromCurrentRoster();
      return;
    }

    if (mode == GenerationMode.excelUpload) {
      await _generateFromExcelUpload();
    }
  }

  Future<void> _generateFromCurrentRoster() async {
    await ref
        .read(scheduleGenerationControllerProvider.notifier)
        .generateFromCurrentRoster(month: 'November');

    if (!mounted) return;

    final currentState = ref.read(scheduleGenerationControllerProvider);
    if (currentState.status == GenerationStatus.success) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Schedule generated from current roster')),
      );
    }
  }

  Future<void> _generateFromExcelUpload() async {
    final file = await showDialog<PlatformFile>(
      context: context,
      builder: (_) => const ExcelUploadDialog(),
    );

    if (!mounted || file == null) return;

    final user = ref.read(currentUserProvider);
    await ref
        .read(scheduleGenerationControllerProvider.notifier)
        .generateFromExcel(
          file,
          targetMonth: '2026-11',
          departmentId: user?.departmentId,
        );

    if (!mounted) return;

    final currentState = ref.read(scheduleGenerationControllerProvider);
    if (currentState.status == GenerationStatus.success) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Schedule generated from ${file.name}')),
      );
    }
  }

  Future<void> _handleExportPressed() async {
    final user = ref.read(currentUserProvider);
    final genState = ref.read(scheduleGenerationControllerProvider);
    final targetMonth = genState.draft?.targetMonth ?? '2026-11';
    final bytes = await ref
        .read(scheduleGenerationControllerProvider.notifier)
        .exportCurrentDraft(
          draftId: genState.draft?.id,
          targetMonth: targetMonth,
          departmentId: user?.departmentId,
        );

    if (!mounted || bytes == null) return;

    final fileName = 'schedule_$targetMonth.xlsx';
    final savedPath = widget.onSaveFile != null
        ? await widget.onSaveFile!(
            fileName: fileName,
            bytes: bytes,
          )
        : await FilePicker.saveFile(
            dialogTitle: 'Save Schedule',
            fileName: fileName,
            bytes: Uint8List.fromList(bytes),
            type: FileType.custom,
            allowedExtensions: const ['xlsx'],
          );

    if (!mounted) return;

    if (savedPath != null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Schedule exported successfully')),
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    ref.listen<ScheduleGenerationState>(
      scheduleGenerationControllerProvider,
      (previous, next) {
        if (next.status == GenerationStatus.error &&
            next.errorMessage != null &&
            next.errorMessage != previous?.errorMessage) {
          ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(
              content: Text(next.errorMessage!),
              backgroundColor: Theme.of(context).colorScheme.error,
            ),
          );
        }
      },
    );

    final genState = ref.watch(scheduleGenerationControllerProvider);

    return Scaffold(
      appBar: AppBar(
        title: const Text('MedShift Admin'),
        leading: Builder(
          builder: (scaffoldContext) => IconButton(
            icon: const CircleAvatar(
              radius: 16,
              child: Icon(Icons.person, size: 20),
            ),
            onPressed: () => Scaffold.of(scaffoldContext).openDrawer(),
          ),
        ),
        actions: [
          IconButton(
            icon: const Icon(Icons.notifications_outlined),
            onPressed: () {},
          ),
        ],
      ),
      drawer: const ProfileDrawer(),
      body: SingleChildScrollView(
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.center,
          children: [
            AdminHeroCard(
              targetMonth: 'November',
              onGeneratePressed: _handleGeneratePressed,
              isGenerating: genState.isSolving,
            ),
            const SizedBox(height: 10),
            const AdminMetricCards(),
            const SizedBox(height: 10),
            DraftPreviewCanvas(
              isGenerated: genState.isGenerated,
              draft: genState.draft,
              isExporting: genState.isExporting,
              onExportPressed: _handleExportPressed,
            ),
          ],
        ),
      ),
      bottomNavigationBar: BottomNavBar(
        currentIndex: _currentIndex,
        showAdminTab: true,
        onTabSelected: (index) => setState(() => _currentIndex = index),
      ),
    );
  }
}
