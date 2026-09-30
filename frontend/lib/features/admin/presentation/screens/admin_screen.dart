import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/admin/presentation/widgets/admin_hero_card.dart';
import 'package:frontend/features/admin/presentation/widgets/admin_metric_cards.dart';
import 'package:frontend/features/admin/presentation/widgets/draft_preview_canvas.dart';
import 'package:frontend/features/admin/presentation/widgets/excel_upload_dialog.dart';
import 'package:frontend/features/admin/presentation/widgets/generation_mode_dialog.dart';
import 'package:frontend/shared/widgets/bottom_nav_bar.dart';
import 'package:frontend/shared/widgets/profile_drawer.dart';

class AdminScreen extends ConsumerStatefulWidget {
  const AdminScreen({super.key});

  @override
  ConsumerState<AdminScreen> createState() => _AdminScreenState();
}

class _AdminScreenState extends ConsumerState<AdminScreen> {
  int _currentIndex = 3;

  Future<void> _handleGeneratePressed() async {
    final mode = await showDialog<GenerationMode>(
      context: context,
      builder: (_) => const GenerationModeDialog(),
    );

    if (!mounted || mode == null) return;

    switch (mode) {
      case (GenerationMode.currentRoster):
        break;
      case (GenerationMode.excelUpload):
        final file = await showDialog<PlatformFile>(
          context: context,
          builder: (_) => const ExcelUploadDialog(),
        );

        if (!mounted || file == null) return;

        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Schedule generated from ${file.name}')),
        );

        break;
    }
  }

  @override
  Widget build(BuildContext context) {
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
            ),
            const SizedBox(height: 10),
            AdminMetricCards(),
            const SizedBox(height: 10),
            DraftPreviewCanvas(),
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
