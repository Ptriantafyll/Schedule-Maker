
  You'll need to add for each feature:

   Layer                                                                     | What to add
  ---------------------------------------------------------------------------|---------------------------------------------------------------------------
   repository                                                                |  soft_delete_X(session, id)  — sets  is_deleted = True
   controller                                                                |  delete_X_controller(id, session)  — validates existence, calls repo
   route                                                                     |  @router.delete("/{id}")  — returns  204 No Content

  ### 2. Status with Logging

  Currently, the backend has no logging configured. It only uses a couple of temporary  print  statements in  main.py .

  #### Where should we add logging?

  1. Application Lifecycle: Log server startup, database initialization, and shutdown (replacing the  print  calls in  main.py ).
  2. API Layer (Routes/Controllers): Log incoming requests, errors, and database transaction issues.
  3. Synchronization (Offline Sync): Since the desktop app will sync data upstream/downstream, logging sync payload sizes, successes, conflicts, and
  network failures is crucial for debugging sync issues.
  4. Optimization Engine (Scheduler): Log solver execution details (e.g., when the CP-SAT solver starts, how many constraints were generated, if it found
  a feasible/optimal solution, or if it timed out).

  #### How should we implement it?

  We can use Python's standard  logging  library. We should:

  1. Define a standard configuration in a new file, e.g.,  src/utils/logging.py , that configures console logging with colors and timestamp formatting.
  2. Initialize this configuration in  src/main.py .
  3. Import  logging  and use  logger = logging.getLogger(__name__)  inside routers, connection setup, and the scheduler.


July 3: Update repository functions for doctor

July 22: The test in test_models.py should go to doctor after moving the models
1. Create the tests for doctor
2. Create the doctor models
3. Update repository functions for doctor
4. Create the controllers

Aug 5:

## Done

1. Created repository test function definitions for doctor - pending implementation

## Todo

1. Create doctor routes definitions
2. Create and implement doctor route tests
3. Create doctor controllers definitions
4. Create and implement doctor controllers tests
5. Create doctor models
6. Create doctor repository definitions (done)
7. Create and implement doctor repository tests
8. Implement doctor routes
9. Implement doctor controllers
10. Implement doctor repository

## Aug 6: 

Done

1. Created doctor routes definitions

Todo

1. Create logger in utils/
2. Create and implement doctor route tests
3. Create doctor controllers definitions
4. Create and implement doctor controllers tests
5. Create doctor models
6. Create doctor repository definitions (done)
7. Create and implement doctor repository tests
8. Implement doctor routes
9.  Implement doctor controllers
10. Implement doctor repository

## Aug 7: 

Done

1. Create logger in utils/

Todo

1. Add logging wherever necessary
2. Create and implement doctor route tests
3. Create doctor controllers definitions
4. Create and implement doctor controllers tests
5. Create doctor models
6. Create doctor repository definitions (done)
7. Create and implement doctor repository tests
8. Implement doctor routes
9.  Implement doctor controllers
10. Implement doctor repository

## Aug 10

Done 

1. Added logging for every request
2. Created doctor route test definitions

Todo:

1. Create doctor controllers definitions
2. Create doctor conterollers test definitions
3. Create doctor models
4. implement doctor repository and tests
5. implement doctor controllers and tests
6. implement doctor routes and tests

# Aug 11

Done:

1. Added route for creating doctor end-to-end with tests

Todo:

1. Do the same with the next routes, don't forget the tests
2. recheck the team and department controllers for error handling

# Aug 12

done

1. Added routes for creating and listing doctor pre assignment

todo: 

1. Do the same with the next routes
2. recheck the team nad department controllers for error handling