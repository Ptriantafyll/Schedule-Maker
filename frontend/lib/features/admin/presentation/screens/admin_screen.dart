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
  })?
  onSaveFile;

  @override
  ConsumerState<AdminScreen> createState() => _AdminScreenState();
}

class _AdminScreenState extends ConsumerState<AdminScreen> {
  int _currentIndex = 3;

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addPostFrameCallback((_) {
      _initializeDashboard();
    });
  }

  void _initializeDashboard() {
    final user = ref.read(currentUserProvider);
    ref
        .read(scheduleGenerationControllerProvider.notifier)
        .initializeDashboard(departmentId: user?.departmentId);
  }

  Future<void> _handlePublishPressed() async {
    final genState = ref.read(scheduleGenerationControllerProvider);
    final targetMonth = genState.selectedMonth;

    final confirmed = await showDialog<bool>(
      context: context,
      builder: (dialogContext) {
        return AlertDialog(
          title: const Text('Publish Schedule'),
          content: Text(
            'Are you sure you want to publish the schedule for ${genState.selectedMonth}? Once published, assignments will be visible to all staff.',
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(dialogContext).pop(false),
              child: const Text('Cancel'),
            ),
            FilledButton(
              onPressed: () => Navigator.of(dialogContext).pop(true),
              child: const Text('Publish'),
            ),
          ],
        );
      },
    );

    if (!mounted || confirmed != true) return;

    final user = ref.read(currentUserProvider);
    await ref
        .read(scheduleGenerationControllerProvider.notifier)
        .publishCurrentDraft(departmentId: user?.departmentId);

    if (!mounted) return;

    final nextState = ref.read(scheduleGenerationControllerProvider);
    if (nextState.status == GenerationStatus.success) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('Schedule for $targetMonth published successfully.'),
        ),
      );
    }
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
    final genState = ref.read(scheduleGenerationControllerProvider);
    await ref
        .read(scheduleGenerationControllerProvider.notifier)
        .generateFromExcel(
          file,
          targetMonth: genState.selectedMonth,
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
    final targetMonth = genState.draft?.targetMonth ?? genState.selectedMonth;
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
        ? await widget.onSaveFile!(fileName: fileName, bytes: bytes)
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
    ref.listen<ScheduleGenerationState>(scheduleGenerationControllerProvider, (
      previous,
      next,
    ) {
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
    });

    final genState = ref.watch(scheduleGenerationControllerProvider);

    final availableMonths = <String>{
      if (genState.targetMonthInfo != null)
        genState.targetMonthInfo!.nextTargetMonth,
      ...genState.scheduleHistory.map((s) => s.targetMonth),
    }.toList();

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
            Padding(
              padding: const EdgeInsets.symmetric(horizontal: 20),
              child: AdminHeroCard(
                targetMonth: genState.selectedMonth,
                isPublished: genState.isPublished,
                availableMonths: availableMonths,
                onMonthSelected: (month) {
                  final user = ref.read(currentUserProvider);
                  ref
                      .read(scheduleGenerationControllerProvider.notifier)
                      .selectMonth(month, departmentId: user?.departmentId);
                },
                onGeneratePressed: _handleGeneratePressed,
                isGenerating: genState.isSolving,
              ),
            ),
            const SizedBox(height: 10),
            const Padding(
              padding: EdgeInsets.symmetric(horizontal: 20),
              child: AdminMetricCards(),
            ),
            const SizedBox(height: 10),
            Padding(
              padding: const EdgeInsets.symmetric(horizontal: 20),
              child: DraftPreviewCanvas(
                isGenerated: genState.isGenerated,
                draft: genState.draft,
                isExporting: genState.isExporting,
                onExportPressed: _handleExportPressed,
                isPublishing: genState.isPublishing,
                onPublishPressed: _handlePublishPressed,
              ),
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
