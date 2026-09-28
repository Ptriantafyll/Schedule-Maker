import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:frontend/features/admin/presentation/widgets/admin_hero_card.dart';
import 'package:frontend/shared/widgets/bottom_nav_bar.dart';
import 'package:frontend/shared/widgets/profile_drawer.dart';

class AdminScreen extends ConsumerStatefulWidget {
  const AdminScreen({super.key});

  @override
  ConsumerState<AdminScreen> createState() => _AdminScreenState();
}

class _AdminScreenState extends ConsumerState<AdminScreen> {
  int _currentIndex = 3;

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
      body: Column(
        crossAxisAlignment: CrossAxisAlignment.center,
        children: [
          AdminHeroCard(targetMonth: 'November', onGeneratePressed: () {}),
        ],
      ),
      bottomNavigationBar: BottomNavBar(
        currentIndex: _currentIndex,
        showAdminTab: true,
        onTabSelected: (index) => setState(() => _currentIndex = index),
      ),
    );
  }
}
